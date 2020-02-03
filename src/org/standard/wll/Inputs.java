package org.standard.wll;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;


/**
 * @author Victoria Torres
 *
 */

public class Inputs {
	public String stand = null;
	public double factor = 1;
	public static double id_dilution;
	int rowindex;
	double diff_2_factor;

	// parameters extracted from the paramater sheet
	String[] raw_data;
	int[] parameter_dilutions;
	String standards;
	String reference_factors;
	String cut_off;
	double correlation_cut_off;
	double slope_cut_off;
	double sloperatio_cut_off;
	
	String [] type_list;
	String RUNcol;
	String IDcol;
	String DILUTIONcol;
	String [] header;

	public String[] get_header() {
		return header;
	}
	public String[] get_types() {
		return type_list;
	}
	public String get_col_RUN() {
		return RUNcol;
	}
	public String get_col_ID() {
		return IDcol;
	}
	public String get_col_DILUTION() {
		return DILUTIONcol;
	}
	
	
	
	public String get_standard() {
		return stand;
	}

	public double get_factor() {
		return factor;
	}

	public double get_id_dilution() {
		return id_dilution;
	}

	public double get_diff_2_factor() {
		return diff_2_factor;
	}

	public String[] get_raw_data() {
		return raw_data;
	}

	public int[] get_dilutions() {
		return parameter_dilutions;
	}

	public String get_standards() {
		return standards;
	}

	public String get_reference_factors() {
		return reference_factors;
	}

	public String get_cut_off() {
		return cut_off;
	}

	public double get_correlation_cut_off() {
		return correlation_cut_off;
	}

	public double get_slope_cut_off() {
		return slope_cut_off;
	}

	public double get_sloperatio_cut_off() {
		return sloperatio_cut_off;
	}

	// This method is used to load the excel files that would be used as inputs for
	// the program. File_name is provided by the user.
	public Workbook load_excel(String file_name) throws IOException, FileNotFoundException, InvalidFormatException {

		// XSSFWorkbook workbook = new XSSFWorkbook(new File(file_name));
		Workbook workbook = WorkbookFactory.create(new File(file_name));
		return workbook;
	}

	public void read_parameters(XSSFSheet parameter_sheet) {
		Row row = parameter_sheet.getRow(1);
		Cell cell = row.getCell(1);
		int num_raw_sheets = (int) cell.getNumericCellValue();
		raw_data = new String[num_raw_sheets];

		for (int i = 0; i < num_raw_sheets; i++) {
			row = parameter_sheet.getRow(2);
			cell = row.getCell((i + 1));
			raw_data[i] = cell.getStringCellValue();
		}

		//rf
		row = parameter_sheet.getRow(3);
		cell = row.getCell(1);
		reference_factors = cell.getStringCellValue();

		//stand
		row = parameter_sheet.getRow(4);
		cell = row.getCell(1);
		standards = cell.getStringCellValue();

		//cut off
		row = parameter_sheet.getRow(5);
		cell = row.getCell(1);
		cut_off = cell.getStringCellValue();

		//correlation cut off
		row = parameter_sheet.getRow(9);
		cell = row.getCell(1);
		correlation_cut_off = cell.getNumericCellValue();

		//slope cut off
		row = parameter_sheet.getRow(10);
		cell = row.getCell(1);
		slope_cut_off = cell.getNumericCellValue();

		//slope ratio cut off
		row = parameter_sheet.getRow(11);
		cell = row.getCell(1);
		sloperatio_cut_off = cell.getNumericCellValue();

		//parameter dilutions
		row = parameter_sheet.getRow(13);
		cell = row.getCell(1);
		int num_dilutions = (int) cell.getNumericCellValue();
		parameter_dilutions = new int[num_dilutions];

		for (int i = 0; i < num_dilutions; i++) {
			row = parameter_sheet.getRow(14);
			cell = row.getCell((i + 1));
			parameter_dilutions[i] = (int) cell.getNumericCellValue();
		}

		//difference factor
		row = parameter_sheet.getRow(16);
		cell = row.getCell(1);
		diff_2_factor = cell.getNumericCellValue();
		
		//virus list
		row = parameter_sheet.getRow(18);
		cell = row.getCell(1);
		int num_types = (int) cell.getNumericCellValue();
		type_list = new String[num_types];

		for (int i = 0; i < num_types; i++) {
			row = parameter_sheet.getRow(19);
			cell = row.getCell((i + 1));
			type_list[i] = cell.getStringCellValue();
		}
		
		//run, id, dilutions columns. Needed for dif
		row = parameter_sheet.getRow(21);
		cell = row.getCell(1);
		RUNcol = cell.getStringCellValue();
		
		row = parameter_sheet.getRow(23);
		cell = row.getCell(1);
		IDcol = cell.getStringCellValue();
		
		row = parameter_sheet.getRow(25);
		cell = row.getCell(1);
		DILUTIONcol = cell.getStringCellValue();
		
		
		row = parameter_sheet.getRow(27);
		cell = row.getCell(1);
		int num_header = (int) cell.getNumericCellValue();
		header = new String[num_header];

		for (int i = 0; i < num_header; i++) {
			row = parameter_sheet.getRow(28);
			cell = row.getCell((i + 1));
			header[i] = cell.getStringCellValue();
		}

	}

	// This method goes through the whole ref factors file and returns an array with
	// ALL the rf.
	public int[] reference_factors(XSSFSheet rf_sheet) {
		int count = count_ints(rf_sheet);
		int[] rf_values = new int[count];
		int i = 0;
		int standard_row;
		int type_column;

		for (Row row : rf_sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					type_column = cell.getColumnIndex();
					standard_row = cell.getRowIndex();
					rf_values[i] = read_reference_factors(rf_sheet, standard_row, type_column);
					i++;
				}
			}
		}
		return rf_values;
	}

	// This method is called by reference_factors to obtain the value in the cell
	// (rf).
	public int read_reference_factors(XSSFSheet rf_sheet, int standard_row, int type_column) {
		int rf = 0;

		XSSFCell rf_value = rf_sheet.getRow(standard_row).getCell(type_column);
		rf = (int) rf_value.getNumericCellValue();

		return rf;
	}

	// gets the number of cells in a document that have ints inside
	// used for reference_factors
	public int count_ints(XSSFSheet rf_sheet) {
		int count = 0;
		for (Row row : rf_sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					count++;
				}
			}
		}
		return count;
	}

	// method used to find the position where the standard is located
	public int find_row(XSSFSheet sheet, String cell_content) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
						return row.getRowNum();
					}
				}
			}
		}
		return 0;
	}

	// methods used to find the position where the type is located
	public int find_column(XSSFSheet sheet, String cell_content) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
						return cell.getColumnIndex();
					}
				}
			}
		}
		return 0;
	}

	// returns a list with all the names of HPVs present in the reference factors
	// sheet
	public String[] hpvlst(XSSFSheet rf_sheet) {
		Row hpvRow = rf_sheet.getRow(0);
		int numCol = hpvRow.getLastCellNum();
		String[] hpv = new String[numCol - 1];
		String content;

		// gets the name of all the HPVs
		for (int j = 1; j <= (numCol - 1); j++) {
			content = hpvRow.getCell(j).getStringCellValue();
			hpv[j - 1] = content;
		}
		return hpv;
	}

	// The HPV list is used as a sort of counter.
	public double id_hpv(XSSFSheet rf_sheet, String hpv) {
		double rf = 0;
		int colindex = find_column(rf_sheet, hpv);

		for (rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {
			Row row = rf_sheet.getRow(rowindex);
			Cell cell = row.getCell(colindex);
			if (cell.getCellType() == CellType.NUMERIC) {
				rf = cell.getNumericCellValue();
				stand = row.getCell(0).getStringCellValue();
			}
		}
		return rf;
	}

	// This method checks the number of dilutions there is by comparing the id and
	// run values
	public int size_dilutionlst(XSSFSheet raw_sheet, int num_dilutions, int pos) {
		// column indexes
		int RUN = find_column(raw_sheet, RUNcol);
		int ID = find_column(raw_sheet, IDcol);

		int size = 1;

		// Initial cells used to then compared with the others
		Row ROW1 = raw_sheet.getRow(pos + 1);
		Cell R1 = ROW1.getCell(RUN);
		Cell ID1 = ROW1.getCell(ID);

		DataFormatter dataformatter = new DataFormatter();
		String r_val;
		String id_val;
		String R1_val = dataformatter.formatCellValue(R1);
		String ID1_val = dataformatter.formatCellValue(ID1);

		rowindex = (pos + 2);
		Row row = raw_sheet.getRow(rowindex);
		Cell r_cell = row.getCell(RUN);
		Cell id_cell = row.getCell(ID);
		r_val = dataformatter.formatCellValue(r_cell);
		id_val = dataformatter.formatCellValue(id_cell);
		int lastrow = raw_sheet.getLastRowNum();
		while (r_val.equals(R1_val) && id_val.equals(ID1_val)) {
			size++;
			if (rowindex <= lastrow-1) {
				row = raw_sheet.getRow(++rowindex);
				r_cell = row.getCell(RUN);
				id_cell = row.getCell(ID);
				r_val = dataformatter.formatCellValue(r_cell);
				id_val = dataformatter.formatCellValue(id_cell);
			}else {break;}
		}
		return size;
	}

	// run and id of the sample needed to present the results
	public String[] run_id(XSSFSheet raw_sheet, int pos) {
		String[] runId = new String[2];
		int RUN = find_column(raw_sheet, RUNcol);
		int ID = find_column(raw_sheet, IDcol);

		DataFormatter dataformatter = new DataFormatter();

		Row row = raw_sheet.getRow(pos + 1);
		Cell r_cell = row.getCell(RUN);
		Cell id_cell = row.getCell(ID);
		runId[0] = dataformatter.formatCellValue(r_cell);
		runId[1] = dataformatter.formatCellValue(id_cell);

		return runId;
	}

	// Gets the values of dilutions
	public double[] get_dilutions(XSSFSheet raw_sheet, int size) {
		int DILUTION = find_column(raw_sheet, DILUTIONcol);
		double[] dilution = new double[size];

		for (int rowindex = 1; rowindex <= size; rowindex++) {
			Row row = raw_sheet.getRow(rowindex);
			Cell d_cell = row.getCell(DILUTION);
			dilution[(rowindex - 1)] = d_cell.getNumericCellValue();
		}
		return dilution;
	}

	// Extracts a line of the raw data document
	public double[] line_raw(XSSFSheet raw_sheet, String type, int pos, int size) {
		int type_col = find_column(raw_sheet, type);
		double[] data = new double[size];

		for (int rowindex = (pos + 1); rowindex < (size + (pos + 1))
				&& rowindex <= raw_sheet.getLastRowNum(); rowindex++) {
			Row row = raw_sheet.getRow(rowindex);
			Cell data_cell = row.getCell(type_col);
			if (data_cell.getCellType() == CellType.NUMERIC) {
				data[((rowindex - (pos + 1)))] = data_cell.getNumericCellValue();
			} else {
				data[(rowindex - (pos + 1))] = 0;
			}
		}
		return data;
	}

	public ArrayList<Double> ctrl_standards(XSSFSheet ctrl_sheet, String type, String[] run_id, double[] dilution) {
		ArrayList<Double> ctrl = new ArrayList<Double>(); //double[] ctrl = new double[size];
		int type_col = find_column(ctrl_sheet, type);

		DataFormatter dataformatter = new DataFormatter();
		String run_c;
		String stand_c;

		for (int rowindex = 1; rowindex <= ctrl_sheet.getLastRowNum(); rowindex++) {
			Row row = ctrl_sheet.getRow(rowindex);
			Cell run_cell = row.getCell(0);
			run_c = dataformatter.formatCellValue(run_cell);
			Cell standard_cell = row.getCell(2);
			stand_c = dataformatter.formatCellValue(standard_cell);

			if (run_c.equals(run_id[0]) && (stand_c.equals(stand))) { // this is needed in case that the id has
																		// different dilutions from the sample data
				Cell dilution_check = row.getCell(3);
				id_dilution = dilution_check.getNumericCellValue();
				if (id_dilution != dilution[0]) {
					factor = (id_dilution / dilution[0]); // first dilution of the id over the starting dil of raw data.
				} // preliminary, need to manipulate later in the Calc class
				while(run_c.equals(run_id[0]) && (stand_c.equals(stand))) {
					Row t_row = ctrl_sheet.getRow(rowindex);
					Cell type_cell = t_row.getCell(type_col);
					ctrl.add(type_cell.getNumericCellValue());
					
					rowindex+= 1;
					row = ctrl_sheet.getRow(rowindex);
					run_cell = row.getCell(0);
					run_c = dataformatter.formatCellValue(run_cell);
					standard_cell = row.getCell(2);
					stand_c = dataformatter.formatCellValue(standard_cell);
				}
				
//				while (r_val.equals(R1_val) && id_val.equals(ID1_val)) {
//					size++;
//					if (rowindex <= lastrow-1) {
//						row = raw_sheet.getRow(++rowindex);
//						r_cell = row.getCell(RUN);
//						id_cell = row.getCell(ID);
//						r_val = dataformatter.formatCellValue(r_cell);
//						id_val = dataformatter.formatCellValue(id_cell);

//				for (int i = 0; i < ctrl.length; i++) {
//					Row t_row = ctrl_sheet.getRow(rowindex + i);
//					Cell type_cell = t_row.getCell(type_col);
//					ctrl[i] = type_cell.getNumericCellValue();
//				}
				break;
			}
		}
		return ctrl;
	}

	// this extracts the value needed for the seropositivity test
	public double cut_off(XSSFSheet cut_sheet, String type) {
		int type_col = find_column(cut_sheet, type);
		Row row = cut_sheet.getRow(1);
		Cell cut = row.getCell(type_col);
		double cut_off = cut.getNumericCellValue();
		return cut_off;
	}

	public boolean seropositivity(double cut_off, double[] data) {
		boolean seropositive = false;
		for (int i = 0; i < data.length; i++) {
			if (data[i] > cut_off) {
				seropositive = true;
			}
		}
		return seropositive;
	}
}
