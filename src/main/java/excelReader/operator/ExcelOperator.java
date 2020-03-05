package excelReader.operator;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

import excelReader.reader.ExcelReader;
import excelReader.writer.ExcelWriter;

public class ExcelOperator {

	public static void main(String[] args) {
		BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
		String s="";
		ExcelWriter xlwriter = null;
		ExcelReader xlreader = null;
		try {
			while (!s.equals("3")) {
			System.out.println("--------------Excel Operator-------------------------------");
			System.out.println("Enter Option to choose from, Numbers only");
			System.out.println("1. Write");
			System.out.println("2. Read");
			s = reader.readLine();
				if (s.equals("1")) {
				System.out.println("Enter Option to choose from");
					System.out.println("1. Add Rows to Sheet");
					s = reader.readLine();
					if (s.equals("1")) {
					System.out.println("Sheet Name Please");
					String sheet = reader.readLine();
					if (xlwriter == null) {
						xlwriter = new ExcelWriter();
							xlwriter.addSheet(sheet);
					}
					} else if (s.equals("2")) {
					System.out.println("Complete Sheet Path");
					String path = reader.readLine();
					System.out.println("Sheet Name Please");
					String sheet = reader.readLine();
					System.out.println("Enter new rows in single Line separated by space");
					String[] rows = reader.readLine().split(" ");
					if (xlwriter == null) {
						xlwriter = new ExcelWriter();
						xlwriter.addRows(sheet, path, rows);
					}
				}
				else {
					System.out.println("Wrong Option Resetting.......");
					continue;
				}
			}
				else if (s.equals("2")) {
					System.out.println("Complete Sheet Path");
					String path = reader.readLine();
					System.out.println("Sheet Name Please");
					String sheet = reader.readLine();
					if (xlreader == null) {
						xlreader = new ExcelReader();
						xlreader.read(sheet, path);
					}
				}
				else if (s.equals("3")) {
					break;
				}
				else {
					System.out.println("Wrong Option Resetting.......");
					continue;
				}
			}
		} catch (NumberFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
