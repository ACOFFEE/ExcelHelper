package com.cf.excel.writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

//import com.bookdao.oasis.utils.uuid.UUIDGenerator;
import com.cf.excel.ExcelConstants;

public abstract class ExcelWriter<T> {
	/**
	 * 每个EXCEL文件的SHEET个数
	 */
	private int sheets = 10;

	/**
	 * 每个SHEET数据量
	 */
	private int rows = 100000;

	/**
	 * 每个cell最大宽度
	 */
	private int cellWithMax = 8000;

	/**
	 * 数据标题
	 */
	public List<String> titles;

	private int textFont = 13;

	public int getRows() {
		return rows;
	}

	public void setRows(int rows) {
		this.rows = rows;
	}

	private Workbook workBook;

	private String fileName;

	private String filePath;

	private int nowWookBookNumber;

	private int dataCount;

	private int stepCount;

	private int officeVersion;

	private String fileSubName;

	private int titleIndex = 0;

	private int defaultWidth;

	private String sheetName = "sheet";

	private File finalFile;

	private String baseDir;

	private int onePageSize = 2000;

	private String excelTitle;

	public int getOnePageSize() {
		return onePageSize;
	}

	public void setOnePageSize(int onePageSize) {
		this.onePageSize = onePageSize;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public void createWork(List<String> titles, String fileName,
			int officeVersion) {
		baseDir = UUID.randomUUID().toString();
		this.titles = titles;
		this.fileName = fileName;
		this.filePath = System.getProperty("java.io.tmpdir");
		this.officeVersion = officeVersion;
		createWorkBook();
		nowWookBookNumber = 1;
	}

	public void createWorkBook() {
		switch (this.officeVersion) {
		case ExcelConstants.OFFICE_2003:
			this.workBook = new HSSFWorkbook();
			this.fileSubName = ".xls";
			break;
		case ExcelConstants.OFFICE_2007:
			this.workBook = new SXSSFWorkbook(1000);
			this.fileSubName = ".xlsx";
			break;
		case ExcelConstants.OFFICE_2010:
			this.workBook = new SXSSFWorkbook(1000);
			this.fileSubName = ".xlsx";
			break;
		}
	}

	public List<Object> rows(T rows) {
		return new ArrayList<Object>();
	}

	public String fileName() {
		switch (this.finalFileType()) {
		case ExcelConstants.FILE_TYPE_ZIP:
			return this.fileName + ".zip";
		default:
			return this.fileName + this.fileSubName;
		}
	}

	/**
	 * 把数据写入Excel
	 * 
	 * @author xuzhen
	 * @param datas
	 *            数据
	 */
	public void pushDataToExcel(List<T> tDatas) {
		if (null == tDatas) {
			throw new NullPointerException("Datas is null!");
		}
		if (null == workBook) {
			createWorkBook();
		}
		pushData(tDatas);
	}

	public synchronized void pushData(List<T> tDatas) {
		int dataSize = tDatas.size();
		int nowSheets = workBook.getNumberOfSheets();
		if (nowSheets > 0) {
			Sheet lastSheet = workBook.getSheetAt(nowSheets - 1);
			int lastSheetRows = lastSheet.getLastRowNum();
			int surplusRows = rows - lastSheetRows;// 剩余可容纳行数
			if (surplusRows > 0) {
				List<T> tempList = null;
				if (dataSize > surplusRows) {
					tempList = tDatas.subList(0, surplusRows);
					pushDataToSheet(lastSheet, tempList);
					tDatas = tDatas.subList(surplusRows, dataSize);
					forEachDatas(tDatas);
				} else {
					pushDataToSheet(lastSheet, tDatas);
				}
			} else {
				forEachDatas(tDatas);
			}
		} else {// 第一次添加数据
			Sheet sheet = null;
			if (dataSize > rows) {
				forEachDatas(tDatas);
			} else {
				sheet = createSheet(sheetName);
				pushDataToSheet(sheet, tDatas);
			}
		}

	}

	private void createFinalFile() {
		StringBuilder pathStr = new StringBuilder();
		pathStr.append(this.filePath).append(File.separator).append(baseDir);
		File fileDir = new File(pathStr.toString());
		if (fileDir.exists() && fileDir.isDirectory()
				&& fileDir.list().length > 1) {
			try {
				finalFile = this.getZipFile();
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			finalFile = new File(pathStr.append(File.separator)
					.append(fileName).append("_").append(nowWookBookNumber)
					.append(this.fileSubName).toString());
		}
	}

	/**
	 * 创建一个新的SHEET,并且添加标题
	 * 
	 * @author xuzhen
	 * @return
	 */
	protected Sheet createSheet(String sheetName) {
		Sheet sheet = null;
		if (null != sheetName && sheetName.trim().length() > 0) {
			sheet = workBook.createSheet(sheetName);
		} else {
			sheet = workBook.createSheet();
		}
		int titleSize = 0;
		CellStyle cellStyle = getRowTitleStyle();
		// 列表标题
		if (null != titles && (titleSize = titles.size()) > 0) {
			// SHEET中添加标题
			Row row = sheet.createRow(titleIndex);
			// 报表标题样式
			for (int i = 0; i < titleSize; i++) {
				Cell cell = row.createCell(i);
				String cellValue = titles.get(i).toString();
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
				// 行宽
				if (0 == defaultWidth) {
					sheet.setColumnWidth(i, cellValue.getBytes().length * 259);
				} else {
					sheet.setColumnWidth(i, defaultWidth);
				}
			}
			// 冻结报表首行（标题行）
			sheet.createFreezePane(0, titleIndex + 1, 0, titleIndex + 1);
		}
		return sheet;
	}

	public void setDefaultWidth(int defaultWidth) {
		this.defaultWidth = defaultWidth;
	}

	public int getTitleIndex() {
		return titleIndex;
	}

	public void setTitleIndex(int titleIndex) {
		this.titleIndex = titleIndex;
	}

	public Font getWorkBookFont() {
		Font font = workBook.createFont();
		font.setFontName("仿宋");
		font.setFontHeightInPoints((short) this.textFont);
		return font;
	}

	/**
	 * 列表标题样式
	 * 
	 * @author zhongcheng
	 */
	private CellStyle getRowTitleStyle() {
		// 报表标题样式
		CellStyle rowTitleStyle = workBook.createCellStyle();
		rowTitleStyle.setFont(getWorkBookFont());
		rowTitleStyle.setBorderLeft(CellStyle.BORDER_THIN);
		rowTitleStyle.setBorderTop(CellStyle.BORDER_THIN);
		rowTitleStyle.setBorderRight(CellStyle.BORDER_THIN);
		rowTitleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		rowTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		rowTitleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		return rowTitleStyle;
	}

	/**
	 * 对数据进行切分，并写入SHEET中
	 * 
	 * @author xuzhen
	 * @param datas
	 *            数据
	 */
	private void forEachDatas(List<T> datas) {
		int dataSize = datas.size();
		int subSize = dataSize / rows;
		for (int i = 0; i < subSize; i++) {
			List<T> tempList = datas.subList(i * rows, (i + 1) * rows);
			Sheet sheet = createSheet(null);
			pushDataToSheet(sheet, tempList);
		}
		if (dataSize % rows > 0) {
			List<T> tempList = datas.subList(subSize * rows, dataSize);
			Sheet sheet = createSheet(null);
			pushDataToSheet(sheet, tempList);
		}
	}

	/**
	 * 把数据写入知道SHEET中
	 * 
	 * @author xuzhen
	 * @param sheet
	 *            SHEET
	 * @param datas
	 *            数据
	 */
	protected void pushDataToSheet(Sheet sheet, List<T> listDatas) {
		int lastRowNumber = sheet.getLastRowNum();
		int dataSize = listDatas.size();
		sheet.autoSizeColumn(1);
		CellStyle cellStyle = workBook.createCellStyle();
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		for (int i = 0; i < dataSize; i++) {
			List<Object> tempDatas = rows(listDatas.get(i));
			Row row = sheet.createRow(null == this.titles ? i + lastRowNumber
					: i + lastRowNumber + 1);
			int tempSize = tempDatas.size();
			for (int j = 0; j < tempSize; j++) {
				Cell cell = row.createCell(j);
				Object value_obj = tempDatas.get(j);
				String cellValue = null == value_obj ? "" : value_obj
						.toString();
				cell.setCellStyle(cellStyle);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(cellValue);
				if (cellValue.getBytes().length * 259 > sheet.getColumnWidth(j)) {
					sheet.setColumnWidth(
							j,
							cellValue.getBytes().length * 259 > cellWithMax ? cellWithMax
									: cellValue.getBytes().length * 259);
				}
			}
		}
		stepCount += dataSize;
		int nowSheets = workBook.getNumberOfSheets();
		// SHEET已经达到最大容量且每个SHEET数据都达到最大容量则写入磁盘
		if (nowSheets == sheets
				&& workBook.getSheetAt(nowSheets - 1).getLastRowNum() >= rows) {
			writeExcelToFile();
			if (this.dataCount > this.stepCount) {
				// workBook = new XSSFWorkbook();
				createWorkBook();
				nowWookBookNumber += 1;
			}
			return;
		}
		// 当数据全部写入Excel对象，则把Excel对象写入磁盘
		if (this.dataCount <= this.stepCount) {
			writeExcelToFile();
		}

	}

	/**
	 * 把一个Excel对象写入磁盘文件
	 * 
	 * @author xuzhen
	 */
	public void writeExcelToFile() {
		StringBuilder pathStr = new StringBuilder();
		pathStr.append(this.filePath).append(File.separator).append(baseDir)
				.append(File.separator).append(fileName).append("_")
				.append(nowWookBookNumber).append(this.fileSubName);
		File excelFile = new File(pathStr.toString());
		File pathFile = excelFile.getParentFile();
		if (!pathFile.exists()) {
			pathFile.mkdirs();
		}

		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(excelFile);
			this.workBook.write(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (null != outputStream) {
				try {
					outputStream.flush();
					outputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	public int excelSize() {
		if (null == finalFile) {
			createFinalFile();
		}
		if (null != finalFile && finalFile.exists()) {
			FileInputStream in = null;
			try {
				in = new FileInputStream(finalFile);
				return in.available();
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (null != in) {
					try {
						in.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return 0;
	}

	public int finalFileType() {
		StringBuilder pathStr = new StringBuilder();
		pathStr.append(this.filePath).append(File.separator).append(baseDir);
		File fileDir = new File(pathStr.toString());
		if (fileDir.exists() && fileDir.isDirectory()
				&& fileDir.list().length > 1) {
			return ExcelConstants.FILE_TYPE_ZIP;
		}
		return ExcelConstants.FILE_TYPE_EXCEL;
	}

	public void writeExcelFile(OutputStream out) {
		if (null == finalFile || !finalFile.exists()) {
			createFinalFile();
		}
		FileInputStream in = null;
		try {
			in = new FileInputStream(finalFile);
			byte[] bytes = new byte[1024];
			int len = -1;
			while ((len = in.read(bytes)) != -1) {
				out.write(bytes, 0, len);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (null != in) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

	}

	public void setDataCount(int dataCount) {
		this.dataCount = dataCount;
	}

	/**
	 * 对生成的Excel文件进行压缩
	 * 
	 * @author xuzhen
	 * @param outFile
	 *            输出的压缩文件
	 * @return
	 * @throws Exception
	 */
	private File getZipFile() throws Exception {
		StringBuilder basePath = new StringBuilder();
		basePath.append(this.filePath).append(File.separator).append(baseDir);
		String outFile = basePath.toString() + ".zip";
		File file = new File(outFile);
		ZipUtils.zip(basePath.toString(), outFile);
		return file;
	}

	/**
	 * 清空临时生成存放EXCEL文件的文件夹
	 * 
	 * @author xuzhen
	 */
	public void clearTempFile() {
		File fileDirs = new File(this.filePath + File.separator + baseDir);
		File files[] = fileDirs.listFiles();
		if (null != files) {
			for (int i = 0; i < files.length; i++) {
				files[i].delete();
			}
		}
		if (null != finalFile) {
			finalFile.delete();
		}
		fileDirs.delete();
	}

	public String getExcelTitle() {
		return excelTitle;
	}

	public void setExcelTitle(String excelTitle) {
		this.excelTitle = excelTitle;
	}

	public InputStream getInputStream() {
		if (null == finalFile || !finalFile.exists()) {
			createFinalFile();
		}
		FileInputStream in = null;
		try {
			in = new FileInputStream(finalFile);
			return in;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
}