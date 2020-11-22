package com.example.demo.service;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.User;
import com.example.demo.repository.FileReadRepository;

@Service
@Transactional
public class ServiceImple implements ReadFileService {

	String JDBC_DRIVER = "com.mysql.jdbc.Driver";
	String DB_URL = "jdbc:mysql://localhost:3306/org";
	String USER = "root";
	String PASS = "password";
	Connection conn = null;
	Statement stmt = null;

	@Autowired
	private FileReadRepository fileReadRepository;

	@Override
	public List<User> findAll() {
		return (List<User>) fileReadRepository.findAll();
	}

	@Override
	public boolean saveDataFromUploadfile(MultipartFile file) {
		boolean isFlag = false;
		String extension = FilenameUtils.getExtension(file.getOriginalFilename());
		if (extension.equalsIgnoreCase("xls") || extension.equalsIgnoreCase("xlsx")) {
			isFlag = readDataFromExcel(file);
		}
		return isFlag;
	}

	private boolean readDataFromExcel(MultipartFile file) {
		Workbook workbook = getWorkBook(file);

		Sheet sheet = workbook.getSheetAt(0);
		createTable(sheet);
		try {
			Iterator<Row> rows = sheet.iterator();
			rows.next();

			conn = DriverManager.getConnection(DB_URL, USER, PASS);
			conn.setAutoCommit(false);

			String sql = "INSERT INTO test1 VALUES (" + getValuesForRow(sheet) + ")";

			PreparedStatement statement = conn.prepareStatement(sql);

			while (rows.hasNext()) {
				Row nextRow = rows.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();
				int temp = 1;

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					if (nextCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						statement.setInt(temp++, (int) nextCell.getNumericCellValue());
					} else {
						statement.setString(temp++, nextCell.getStringCellValue());
					}
				}
				// statement.addBatch();
				// if (count % batchSize == 0) {
				statement.executeBatch();
				// }

			}
			statement.executeBatch();

			conn.commit();
			conn.close();
		} catch (SQLException ex) {
			System.out.println("Database error.......");
			ex.printStackTrace();
		}

		return true;
	}

	private String getValuesForRow(Sheet sheet) {
		String temp = "";
		int totalColumnCount = sheet.getRow(0).getLastCellNum();
		for (int i = 0; i < totalColumnCount - 1; i++) {
			temp += "?,";
		}
		return temp + "?";
	}

	private void createTable(Sheet sheet) {
		Row firstRow = sheet.getRow(1);
		int[] arr = new int[firstRow.getLastCellNum()];
		for (int i = 0; i < firstRow.getLastCellNum(); i++) {
			arr[i] = firstRow.getCell(i).getCellType();
		}

		try {

			Class.forName("com.mysql.jdbc.Driver");
			conn = DriverManager.getConnection(DB_URL, USER, PASS);
			System.out.println("Connected database successfully...");
			stmt = conn.createStatement();

			String sql = "CREATE TABLE IF NOT EXISTS test1 " + "(id INTEGER not NULL, " + getColumns(sheet, arr)
					+ " PRIMARY KEY ( id ))";

			stmt.executeUpdate(sql);
			System.out.println("Created table in given database...");
		} catch (SQLException se) {
			se.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (stmt != null)
					conn.close();
			} catch (SQLException se) {
			}
			try {
				if (conn != null)
					conn.close();
			} catch (SQLException se) {
				se.printStackTrace();
			} // end finally try
		} // end try

	}

	private String getColumns(Sheet sheet, int[] arr) {
		String sql = "";
		Row headerRow = sheet.getRow(0);
		Row headerRowValue = sheet.getRow(1);
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			sql += headerRow.getCell(i) + " " + checkCellType(headerRowValue.getCell(i).getCellType()) + "";
		}
		return sql;
	}

	private String checkCellType(int cellType) {
		String sql = "";
		if (cellType == 1) {
			sql = " VARCHAR(255), ";
			return sql;
		} else if (cellType == 0) {
			sql = " INTEGER, ";
			return sql;
		} else {
			return " VARCHAR(255), ";
		}
	}

	private Workbook getWorkBook(MultipartFile file) {

		Workbook workbook = null;
		String extension = FilenameUtils.getExtension(file.getOriginalFilename());
		try {
			if (extension.equalsIgnoreCase("xlsx")) {
				workbook = new XSSFWorkbook(file.getInputStream());
			} else if (extension.equalsIgnoreCase("xls")) {
				workbook = new HSSFWorkbook(file.getInputStream());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return workbook;
	}

}
