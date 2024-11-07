package main;

public class Sourcefile {
	private String fileName;
	private Filetype fileExtension;
	private String filePath;

	public Sourcefile(String fileName, Filetype fileExtension, String filePath) {
		this.fileName = fileName;
		this.fileExtension = fileExtension;
		this.filePath = filePath;
	}

	public enum Filetype {
		xlsx,
		xls,
		docx,
		doc
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public Filetype getFileExtension() {
		return fileExtension;
	}

	public void setFileExtension(Filetype fileExtension) {
		this.fileExtension = fileExtension;
	}

	public String getFilePath() {
		return filePath;
	}

	public void setFilePath(String fileLocation) {
		this.filePath = fileLocation;
	}

	@Override
	public String toString() {
		return getFileName();
	}

}
