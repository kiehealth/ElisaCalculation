package org.standard.wll;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Victoria Torres
 *
 */

public class Outputs {
	String[] header;

	// Creates a new sheet named after the raw data page name, virus type, and
	// standard. Creates the header.
	public XSSFSheet create_sheet(XSSFWorkbook workbook, String month, String type, String standard,
			double[] parameter_dilutions) throws IOException {
		String name = (month + "  " + type + "  " + standard);
		XSSFSheet sheet = workbook.createSheet(name);

		CellStyle header_style = header_cellstyle(workbook);

		// header information
		String[] tocopy = { "wPLL", "rfl", "PLL", "correlation", "slope", "slope ratio", "sero+", "Error", "Comment" };

		header = new String[parameter_dilutions.length + 11]; // 11 = run, id, results(3), stats(3), seropositivity,
																// errors(2),

		header[0] = "Run";
		header[1] = "id";
		for (int i = 0; i < parameter_dilutions.length; i++) {
			header[i + 2] = String.valueOf(parameter_dilutions[i]);
		}
		int index = 0;
		for (int i = (2 + parameter_dilutions.length); i < header.length; i++) {
			header[i] = tocopy[index++];
		}

		Row row = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(header[i]);
			cell.setCellStyle(header_style);
		}
		return sheet;
	}

	// Puts the data and results in the same array.
	public double[] data_results(double dilution, double[] parameter_dilutions, double[] data, double wPLL, double rfl,
			double pll, double correlation, double slope, double slope_ratio) {
		int num_of_dilutions = data.length; // num of dilutions for this standard
		double[] result = new double[header.length - 5]; // everything except the run, id, sero+, and errors(2)
		int start = 0;
		int index = 0;
		int i = 0;

		try {
		// fixes the data to be placed underneath the right column
		while (dilution != parameter_dilutions[i]) {
			i++;
		}
		} catch (NullPointerException e) {
			System.out.println("You have a dilution not specified in the parameters");
		}
		for (start = i; start < (num_of_dilutions + i); start++) {
			result[start] = data[index];
			index++;
		}

		result[result.length - 6] = wPLL;
		result[result.length - 5] = rfl;
		result[result.length - 4] = pll;
		result[result.length - 3] = correlation;
		result[result.length - 2] = slope;
		result[result.length - 1] = slope_ratio;

		return result;
	}

	// Used for the control and standards lines.
	public void write_data(CellStyle style, CellStyle warning_style, XSSFSheet sheet, int index, String[] run_id,
			double[] data_results, double correlation_cut_off, double slope_cut_off, double sloperatio_cut_off) {
		Row row = sheet.createRow(index);
		Cell cell;
		Boolean error = false;
		for (int i = 0; i < run_id.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(run_id[i]);
			cell.setCellStyle(style);
		}

		int size = (data_results.length + 2);

		for (int i = 2; i < (size - 3); i++) {
			cell = row.createCell(i);
			if (data_results[i - 2] == 0) {
				// this sets cells to blank
			} else {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(style);
			}
		}

		cell = row.createCell(size - 3);
		if (data_results[data_results.length - 3] < correlation_cut_off) {
			cell.setCellValue(data_results[data_results.length - 3]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[data_results.length - 3]);
			cell.setCellStyle(style);
		}

		cell = row.createCell(size - 2);
		if (data_results[data_results.length - 2] > slope_cut_off) {
			cell.setCellValue(data_results[data_results.length - 2]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[data_results.length - 2]);
			cell.setCellStyle(style);
		}

		cell = row.createCell(size - 1);
		if (data_results[data_results.length - 1] < sloperatio_cut_off) { // you can do this a method take abs
			cell.setCellValue(data_results[data_results.length - 1]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[data_results.length - 1]);
			cell.setCellStyle(style);
		}

		// seropositive
		cell = row.createCell(size);
		cell.setCellValue(1);
		cell.setCellStyle(style);

		// errors, if any of the error is true then you have an error type 1
		if (error) {
			cell = row.createCell(header.length - 2);
			cell.setCellValue(1);
			cell.setCellStyle(warning_style);
			cell = row.createCell(header.length - 1);
			cell.setCellValue("Poor Standard");
			error = false;
		}

	}

	// Used for seropositive samples.
	public void sswrite_data(CellStyle style, CellStyle warning_style, XSSFSheet sheet, int index, String[] run_id,
			double[] parameter_dilutions, double[] data_results, double[] data_calculations, double correlation_cut_off,
			double slope_cut_off, double sloperatio_cut_off) {

		Row row = sheet.createRow(index);
		Cell cell;
		boolean error = false;
		int error_check = 0;

		for (int i = 0; i < run_id.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(run_id[i]);
			cell.setCellStyle(style);
		}

		int raw_size = (parameter_dilutions.length + 2);
		int size = data_results.length;

		// prints the raw data
		for (int i = 2; i < raw_size; i++) {
			cell = row.createCell(i);
			if (data_results[i - 2] == 0) {
				// this sets cells to blank
			} else if (i-2 < data_calculations.length && data_results[i - 2] != data_calculations[i - 2]) {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(warning_style);
				error_check++;
			} else {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(style);
			}
		}

		// prints the results wPLL, rfl, PLL
		for (int i = raw_size; i < (raw_size + 3); i++) {
			cell = row.createCell(i);
			if (data_results[i - 2] == 0) {
				// this sets cells to blank
			} else {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(style);
			}
		}

		cell = row.createCell(raw_size + 3);
		if (data_results[size - 3] < correlation_cut_off) {
			cell.setCellValue(data_results[size - 3]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[size - 3]);
			cell.setCellStyle(style);
		}

		cell = row.createCell(raw_size + 4);
		if (data_results[size - 2] > slope_cut_off) {
			cell.setCellValue(data_results[size - 2]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[size - 2]);
			cell.setCellStyle(style);
		}

		cell = row.createCell(raw_size + 5);
		if (data_results[size - 1] < sloperatio_cut_off) { // you can do this a method take abs
			cell.setCellValue(data_results[size - 1]);
			cell.setCellStyle(warning_style);
			error = true;
		} else {
			cell.setCellValue(data_results[size - 1]);
			cell.setCellStyle(style);
		}

		// seropositive
		cell = row.createCell(raw_size + 6);
		cell.setCellValue(1);
		cell.setCellStyle(style);

		// errors
		if (data_calculations.length == error_check || ((data_calculations.length - error_check) == 1)) {
			cell = row.createCell(raw_size + 7);
			cell.setCellValue(2);
			cell.setCellStyle(warning_style);
			cell = row.createCell(raw_size + 8);
			cell.setCellValue("Must be re-tested");
			for (int colindex = raw_size; colindex < (raw_size + 6); colindex++) {
				row.removeCell(row.getCell(colindex));
			}

		} else if (error) {
			cell = row.createCell(raw_size + 7);
			cell.setCellValue(3);
			cell.setCellStyle(warning_style);
			cell = row.createCell(raw_size + 8);
			cell.setCellValue("Could be re-tested");
			error = false;
		} else if ((data_calculations.length - error_check) == 2) {
			cell = row.createCell(raw_size + 7);
			cell.setCellValue(4);
			cell.setCellStyle(warning_style);
			cell = row.createCell(raw_size + 8);
			cell.setCellValue("Two dilutions used");
		}

	}

	// Used for seronegative samples
	public void swrite_data(CellStyle style, XSSFSheet sheet, int index, String[] run_id,
			double[] data_results, double correlation_cut_off, double slope_cut_off, double sloperatio_cut_off) {
		Row row = sheet.createRow(index);
		Cell cell;
		for (int i = 0; i < run_id.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(run_id[i]);
			cell.setCellStyle(style);
		}

		int size = (data_results.length + 2);

		for (int i = 2; i < (size - 3); i++) {
			cell = row.createCell(i);
			if (data_results[i - 2] == 0) {
				// this sets cells to blank
			} else {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(style);
			}
		}

		// seronegative
		cell = row.createCell(size);
		cell.setCellValue(0);
		cell.setCellStyle(style);


	}

	public CellStyle header_cellstyle(XSSFWorkbook workbook) {
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.BLACK.getIndex());
		CellStyle header_style = workbook.createCellStyle();
		header_style.setFont(headerFont);
		header_style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		header_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		header_style.setBorderTop(BorderStyle.THIN);
		header_style.setBorderBottom(BorderStyle.THIN);
		header_style.setBorderLeft(BorderStyle.THIN);
		header_style.setBorderRight(BorderStyle.THIN);
		return header_style;
	}

	public CellStyle seronegative_cellstyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		return style;
	}

	public CellStyle default_cellstyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		return style;
	}

	public CellStyle warning_cellstyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		return style;
	}

	// writes the output to a file
	public void output_file(XSSFWorkbook workbook, XSSFWorkbook input, String output_name, String output_path)
			throws IOException {
		input.close();
		FileOutputStream file_out = new FileOutputStream( output_path + output_name + ".xlsx");
		workbook.write(file_out);
		file_out.close();
		workbook.close();
	}

}