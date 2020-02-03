package org.standard.wll;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Victoria Torres
 *
 */

public class Main {

	public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {

		// Used to get the users input
		Inputs ip = new Inputs();
		Outputs op = new Outputs();
		Calculations cp = new Calculations();

		// help the user out to fill all the sections
		if (args.length < 3) {
			System.out.println(" \n Please re-run your program like this: \n "
					+ "java program.jar /path/to/input name_of_output /path/to/output");
			System.exit(-1);
		}
		String input_path = args[0];
		String output_name = args[1];
		String output_path = args[2];

		// loads the input excelfile
		XSSFWorkbook input = (XSSFWorkbook) ip.load_excel(input_path);

		// all the parameters of the program will be read now
		XSSFSheet parameter_sheet = input.getSheet("Parameters");
		ip.read_parameters(parameter_sheet);

		// Assigns names to the sheets
		XSSFSheet ctrl_sheet = input.getSheet(ip.get_standards());
		XSSFSheet rf_sheet = input.getSheet(ip.get_reference_factors());
		XSSFSheet cutoff_sheet = input.getSheet(ip.get_cut_off());

		String[] hpv = ip.hpvlst(rf_sheet); // creates the list of all hpvs
		int max = ip.count_ints(rf_sheet); // counts how many different rf there is to do

		XSSFWorkbook output = new XSSFWorkbook();

		CellStyle error_cellstyle = op.seronegative_cellstyle(output);
		CellStyle default_cellstyle = op.default_cellstyle(output);
		CellStyle warning_cellstyle = op.warning_cellstyle(output);

		String[] raw_data = ip.get_raw_data();

		// counter for all the raw data sheets
		for (int raw_counter = 0; raw_counter < raw_data.length; raw_counter++) {
			XSSFSheet raw_sheet = input.getSheet(raw_data[raw_counter]);

			// counter for viruses
			for (int master_counter = 0; master_counter < max; master_counter++) {

				double rf = ip.id_hpv(rf_sheet, hpv[master_counter]); // rf factor and id for that specific virus
				double cut_off_value = ip.cut_off(cutoff_sheet, hpv[master_counter]);

				// This is all the processing required for a raw data sheet
				int size = ip.size_dilutionlst(raw_sheet, ip.get_dilutions().length, 0);
				double[] data = new double[size];
				double[] data_calculations = new double[size];

				double[] dilution = ip.get_dilutions(raw_sheet, size); // gets the dilution list
				ArrayList<Double> ctrl; // will be done in the while loop

				int[] parameter_dilutions = ip.get_dilutions();

				XSSFSheet out_sheet = op.create_sheet(output, raw_data[raw_counter], hpv[master_counter],
						ip.get_standard(), parameter_dilutions);

				int df = cp.calculate_df(dilution);

				int index = 1; // this is the index on the row of the output sheet
				double[] log = new double[size];
				double[] log_ctrl = new double[size];
				double wPLL_slope;
				double pll_slope;
				double slope;
				double meanX;
				double meanY;
				double wPLL;
				double rfl;
				double pll;
				double correlation;
				double slope_ratio;
				double[] data_results;
				double[] fixed_data;

				double ctrl_Ymean = 0;
				double ctrl_Xmean = 0;
				double SXX = 0;
				double SXY = 0;
				double first_slope = 0;
				double rfl_denominator = 0;

				// to check if this is a new run, if it is
				// get a new value from ctrls & standards
				String run_check = "0";
				boolean first_line = false;
				String id_check = " ";

				int pos = 0;

				System.out.println("processing");
		//		int lastrow = raw_sheet.getPhysicalNumberOfRows(); 
				while (pos < raw_sheet.getLastRowNum()) { 
					
					String[] run_id = ip.run_id(raw_sheet, pos);
					if (!run_id[0].equals(run_check)) {  //Checks if we are looking at different run, to get new ctrl
						first_line = true;
						size = ip.size_dilutionlst(raw_sheet, ip.get_dilutions().length, pos);
						data = new double[size];
						dilution = ip.get_dilutions(raw_sheet, size);
					}
					if(!run_id[1].equals(id_check)) { //accounts for different amount of dilutions in the same run
						size = ip.size_dilutionlst(raw_sheet, ip.get_dilutions().length, pos);
						data = new double[size];
						dilution = ip.get_dilutions(raw_sheet, size);
					}
					id_check = run_id[1];
					run_check = run_id[0];

					// control and standards line
					if (first_line) {
						
						ctrl = ip.ctrl_standards(ctrl_sheet, hpv[master_counter], run_id, dilution);
						
						
						log_ctrl = cp.log_resultsCTRL(ctrl);
						ctrl_Ymean = cp.Ymean(log_ctrl);
						ctrl_Xmean = cp.Xmean(log_ctrl);
						SXX = cp.sxx(log_ctrl, ctrl_Xmean);
						SXY = cp.sxy(log_ctrl, ctrl_Xmean, ctrl_Ymean);
						first_slope = (SXY / SXX);
						rfl_denominator = (ctrl_Xmean - (ctrl_Ymean / first_slope));

						// first line
						wPLL_slope = cp.slopewPLL(log, ctrl_Xmean, ctrl_Ymean, SXX, SXY);
						wPLL = cp.wPLL(rf, df, wPLL_slope, ctrl_Xmean, ctrl_Ymean, ctrl_Xmean, ctrl_Ymean);
						rfl = cp.rfl(rf, df, rfl_denominator, first_slope, ctrl_Xmean, ctrl_Ymean);
						pll = cp.pll(rf, df, first_slope, ctrl_Xmean, ctrl_Ymean, ctrl_Xmean, ctrl_Ymean);
						slope_ratio = (first_slope / first_slope);
						correlation = cp.correlation(log_ctrl);
						data_results = op.data_resultsCTRL(ip.get_id_dilution(), parameter_dilutions, ctrl, wPLL, rfl, pll,
								correlation, first_slope, slope_ratio);
						String temp = run_id[1];
						run_id[1] = ip.get_standard();
						op.write_data(default_cellstyle, warning_cellstyle, out_sheet, index, run_id, data_results,
								ip.get_correlation_cut_off(), ip.get_slope_cut_off(), ip.get_sloperatio_cut_off());
						run_id[1] = temp;
						index++;
						first_line = false;
					}

					data = ip.line_raw(raw_sheet, hpv[master_counter], pos, size); // data = Inputs.line_raw(raw_sheet,
																					// "HPV 6", pos, size);
					boolean seropositive = ip.seropositivity(cut_off_value, data);

					if (!seropositive) {
						op.swrite_data(error_cellstyle, out_sheet, index, run_id, data);
						index++;
					}

					// only seropositive samples should be used for calculations
					if (seropositive) {

						// removing values to get the negative slope
						double id_dil = ip.get_id_dilution();

						data_calculations = cp.fix_negative_slope(data, ip.get_dilutions(), ip.get_diff_2_factor(),
								id_dil);
						fixed_data = cp.fix_array(data_calculations);
						log = cp.log_results(fixed_data);

						// calculations for lines other than the reference (first line)
						meanX = cp.Xmean(log);
						meanY = cp.Ymean(log);
						wPLL_slope = cp.slopewPLL(log, meanX, meanY, SXX, SXY);
						slope = cp.slope(log, meanX, meanY);
						pll_slope = ((first_slope + slope) / 2);

						// to write
						double factor = cp.get_factor();

						wPLL = cp.wPLL(rf, df, wPLL_slope, meanX, meanY, ctrl_Xmean, ctrl_Ymean);
						wPLL = (wPLL / factor);
						rfl = cp.rfl(rf, df, rfl_denominator, first_slope, meanX, meanY);
						rfl = (rfl / factor);
						pll = cp.pll(rf, df, pll_slope, meanX, meanY, ctrl_Xmean, ctrl_Ymean);
						pll = (pll / factor);
						slope_ratio = (slope / first_slope);
						correlation = cp.correlation(log);

						double d = dilution[0];
						data_results = op.data_results(d, parameter_dilutions, data, wPLL, rfl, pll, correlation, slope,
								slope_ratio);
						op.sswrite_data(default_cellstyle, warning_cellstyle, out_sheet, index, run_id,
								parameter_dilutions, data_results, data_calculations, ip.get_correlation_cut_off(),
								ip.get_slope_cut_off(), ip.get_sloperatio_cut_off());
						index++;
					}
					pos += size; // counter used for extracting the next line
				}
				// inside for loop but outside while loop
			}
			// this is outside the virus loop

		}
		// this is outside the raw data loop

		op.output_file(output, input, output_name, output_path);
		System.out.println("file created");
	}

}
