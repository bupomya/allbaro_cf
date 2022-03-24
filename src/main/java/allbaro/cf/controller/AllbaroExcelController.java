package allbaro.cf.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import allbaro.cf.dto.CF_dto;

public class AllbaroExcelController {

	public void outputExcel(ArrayList<CF_dto> data) throws FileNotFoundException, IOException {

		String path = "C:/Users/chox6/OneDrive/바탕 화면//새 폴더";
		String fileName = "test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("dummy");
		XSSFRow curRow = sheet.createRow(0);
		XSSFCell cell = curRow.createCell(7);

		writeHeader("2021년", "10월", sheet, curRow, cell);
		writeDetail(sheet, curRow, cell, data);

		workbook.write(new FileOutputStream(path + "/" + fileName));
		workbook = new XSSFWorkbook(path + "/" + fileName);
	}

	// 실제 계산값 작성
	public void writeDetail(XSSFSheet sheet, XSSFRow curRow, XSSFCell cell, ArrayList<CF_dto> bData) {
		/*
		 * 1. 수집,재활용을 기준으로 data를 2개의 컬렉션으로 구분 2. 인계서일련번호를 기준으로 날짜가 같으면 재활용항목까지 모두 작성
		 * 2-1. 같은날 재활용하지 않았을 시, 재활용컬렉션에 있는 재활용일자 확인 후 작성 3. 파일로 출력
		 */
		ArrayList<CF_dto> inData = new ArrayList<>();
		ArrayList<CF_dto> outData = new ArrayList<>();
		// 0:수집 1:재활용
		inData = getInputData(bData, 0);	//수집 데이터
		outData = getInputData(bData, 1);	//재활용 데이터
		int rowNo = 6;
		
		for (int cnt = 0; cnt < inData.size(); cnt++) {
			short cellNo = 0;
			
			for(int outCnt=0; outCnt<outData.size();outCnt++) {
				if(inData.get(cnt).getCode()==outData.get(outCnt).getCode()) {//인계서번호로 확인
					if(inData.get(cnt).getDate()==outData.get(outCnt).getDate()) {//날짜 확인
			curRow = sheet.createRow(rowNo);
			// 날짜작성
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(inData.get(cnt).getDate());

			// 업소명
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(inData.get(cnt).getCompanyName());

			// 수집량
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(inData.get(cnt).getInAmount());

			// 폐기물재활용량
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(inData.get(cnt).getInAmount());

			// 종류
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(inData.get(cnt).getKinds());

			// 생산량 = 수집량-수탁물폐기물보관량
			cell = curRow.createCell(cellNo++);
			cell.setCellFormula("C" + (rowNo + 1) + "-" + "G" + (rowNo + 1));

			// 수탁물폐기물보관량
			cell = curRow.createCell(cellNo++);
			cell.setCellValue(1);

			// 상호
			cell = curRow.createCell(cellNo++);
			cell.setCellValue("C.F");

			// 소재지
			cell = curRow.createCell(cellNo++);
			sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, 8, 9));
			cell.setCellValue("죽산면 걸미로 478-12");

			// 사용용도
			cell = curRow.createCell(++cellNo);
			cell.setCellValue("재활용");
			
			// 공급량
			cell = curRow.createCell(++cellNo);
			
					if(inData.get(cnt).getInAmount()==outData.get(outCnt).getOutAmount()) {//수량확인
					cell.setCellFormula(outData.get(outCnt).getOutAmount() + "-" + "G" + (rowNo + 1));
						break;
					}else { //보관량
						cell.setCellFormula(outData.get(outCnt).getOutAmount() + "-" + "G" + (rowNo + 1));
						
						cell = curRow.createCell(++cellNo);
						
						cell.setCellFormula(inData.get(cnt).getInAmount() + "-" + outData.get(outCnt).getOutAmount());
						break;
					}
				}
			}
			}// end for statement
			
			rowNo++;
		}
	}

	public ArrayList<CF_dto> getInputData(ArrayList<CF_dto> bData, int flag) {
		// flag로 수집날짜,재활용날짜 data구분 0:수집 1:재활용
		ArrayList<CF_dto> data = new ArrayList<CF_dto>();// 수집,재활용 dto 구별

		switch (flag) {
		case 0:
			for (int i = 0; i < bData.size(); i++) {
				if (bData.get(i).getOutAmount() == 0) {
					data.add(bData.get(i));
				}
			}
			break;

		case 1:
			for (int i = 0; i < bData.size(); i++) {
				if (bData.get(i).getInAmount() == 0) {
					data.add(bData.get(i));
				}
			}

		}

		for (int i = 0; i < data.size(); i++) {
			System.out.println(data.get(i).toString());
		}

		return data;
	}

	// make dummy data
	public ArrayList<CF_dto> makeDummy() {
		ArrayList<CF_dto> data = new ArrayList<CF_dto>();
		// dummy data
		data.add(new CF_dto("2021-10-12", "종류1", "123123", "112233445", "A company", 50, 0, "11223344"));
		data.add(new CF_dto("2021-10-12", "종류1", "123123", "112233445", "A company", 0, 30, "11223344"));
		data.add(new CF_dto("2021-10-14", "종류2", "321321", "123456789", "A company", 2000, 0, "44332211"));
		data.add(new CF_dto("2021-10-14", "종류2", "321321", "123456789", "B company", 0, 2000, "44332211"));
		data.add(new CF_dto("2021-10-15", "종류1", "123123", "112233445", "A company", 30, 0, "12312312"));
		data.add(new CF_dto("2021-10-15", "종류1", "123123", "112233445", "A company", 0, 30, "12312312"));
		data.add(new CF_dto("2021-10-19", "종류1", "123123", "112233445", "A company", 90, 0, "32132132"));
		data.add(new CF_dto("2021-10-19", "종류1", "123123", "112233445", "A company", 0, 90, "32132132"));
		data.add(new CF_dto("2021-10-19", "종류1", "123123", "987654321", "D company", 900, 0, "32132112"));
		data.add(new CF_dto("2021-10-19", "종류1", "123123", "987654321", "D company", 0, 900, "32132112"));
		data.add(new CF_dto("2021-10-20", "종류2", "321321", "123456789", "B company", 8000, 0, "32132155"));
		data.add(new CF_dto("2021-10-20", "종류2", "321321", "123456789", "B company", 0, 8000, "32132155"));
		data.add(new CF_dto("2021-10-25", "종류1", "123123", "112233445", "A company", 100, 0, "32132166"));
		data.add(new CF_dto("2021-10-26", "종류3", "112233", "111111111", "C company", 200, 0, "32132177"));
		data.add(new CF_dto("2021-10-26", "종류3", "112233", "111111111", "C company", 0, 200, "32132177"));
		data.add(new CF_dto("2021-10-26", "종류1", "123123", "112233445", "A company", 0, 100, "32132166"));
		data.add(new CF_dto("2021-10-26", "종류1", "123123", "987654321", "D company", 500, 0, "32132188"));
		data.add(new CF_dto("2021-10-26", "종류1", "123123", "987654321", "D company", 0, 500, "32132188"));
		data.add(new CF_dto("2021-10-26", "종류4", "132132", "111111111", "C company", 600, 0, "32132199"));
		data.add(new CF_dto("2021-10-26", "종류4", "132132", "111111111", "C company", 0, 600, "32132199"));
		data.add(new CF_dto("2021-10-26", "종류2", "321321", "123456789", "B company", 3000, 0, "32132100"));
		data.add(new CF_dto("2021-10-27", "종류2", "321321", "123456789", "B company", 0, 3000, "32132100"));
		data.add(new CF_dto("2021-10-29", "종류1", "123123", "112233445", "A company", 50, 0, "11223344"));
		data.add(new CF_dto("2021-10-29", "종류1", "123123", "112233445", "A company", 0, 50, "11223344"));

		return data;
	}

	// 경로에 있는 엑셀 파일 읽어오기

	public void writeHeader(String year, String month, XSSFSheet xssfSheet, XSSFRow xssfRow, XSSFCell xssfCell)
			throws FileNotFoundException, IOException {

		int rowNo = 0;

		xssfRow = xssfSheet.createRow(rowNo++); // 행 객체 추가
		xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
		xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8)); // 첫행, 마지막행,첫열,마지막열
		// 타이틀 생성
		xssfCell.setCellValue("[별지 제45호서식] <개정 2008.8.4>");

		// (앞쪽)
		xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 13));

		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 0);
		xssfSheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 13));
//		xssfCell = xssfRow.getCell(1);
		xssfCell.setCellValue("폐기물 수탁 재활용 관리대장");

		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 0);
		xssfSheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 13));
		xssfCell.setCellValue("(재활용대상 폐기물의 종류:)(단위: 톤)");

		// year
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 0);
		xssfCell.setCellValue(year);

		xssfCell = xssfRow.createCell((short) 1);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 2));
		xssfCell.setCellValue("수집ㆍ운반내용");

		xssfCell = xssfRow.createCell((short) 3);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 5));
		xssfCell.setCellValue("폐기물재활용내용");

		xssfCell = xssfRow.createCell((short) 6);
		xssfCell.setCellValue("수탁"); // "수탁" 폐기물 보관량

		xssfCell = xssfRow.createCell((short) 7);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 3, 7, 12));
		xssfCell.setCellValue("재활용제품의 공급 및 보관내용");

		xssfCell = xssfRow.createCell((short) 13);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 5, 13, 13));
		xssfCell.setCellValue("결재");

		// month
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 0);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 5, 0, 0));
		xssfCell.setCellValue(month);

		xssfCell = xssfRow.createCell((short) 1);
		xssfCell.setCellValue("수집");

		xssfCell = xssfRow.createCell((short) 2);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 5, 2, 2));
		xssfCell.setCellValue("수집량");

		xssfCell = xssfRow.createCell((short) 3);
		xssfCell.setCellValue("폐기물");// "폐기물" 재활용량

		xssfCell = xssfRow.createCell((short) 4);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 4, 4, 5));
		xssfCell.setCellValue("재활용제품");

		xssfCell = xssfRow.createCell((short) 6);
		xssfCell.setCellValue("폐기물");// 수탁 "폐기물" 보관량

		xssfCell = xssfRow.createCell((short) 7);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 4, 7, 10));
		xssfCell.setCellValue("공급처");

		xssfCell = xssfRow.createCell((short) 11);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 5, 11, 11));
		xssfCell.setCellValue("공급량");

		xssfCell = xssfRow.createCell((short) 12);
		xssfSheet.addMergedRegion(new CellRangeAddress(4, 5, 12, 12));
		xssfCell.setCellValue("보관량");

		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 1);
		xssfCell.setCellValue("업소명");

		xssfCell = xssfRow.createCell((short) 3);
		xssfCell.setCellValue("재활용량");// 폐기물 "재활용량"

		xssfCell = xssfRow.createCell((short) 4);
		xssfCell.setCellValue("종류");

		xssfCell = xssfRow.createCell((short) 5);
		xssfCell.setCellValue("생산량");

		xssfCell = xssfRow.createCell((short) 6);
		xssfCell.setCellValue("보관량");// 수탁 폐기물 "보관량"

		xssfCell = xssfRow.createCell((short) 7);
		xssfCell.setCellValue("상호(대표자)");

		xssfCell = xssfRow.createCell((short) 8);
		xssfSheet.addMergedRegion(new CellRangeAddress(5, 5, 8, 9));
		xssfCell.setCellValue("소재지");

		xssfCell = xssfRow.createCell((short) 10);
		xssfCell.setCellValue("사용용도");

	}// writeHeader

	public ArrayList<CF_dto> getData(FileInputStream file) {
		ArrayList<CF_dto> data = new ArrayList<CF_dto>();
		try {
//			XSSFWorkbook workbook = new XSSFWorkbook(file);

//			XSSFSheet sheet = workbook.getSheetAt(0); 
//			int rows = sheet.getPhysicalNumberOfRows();

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					// Check the cell type and format accordingly
					switch (cell.getCellType())

					{
					case NUMERIC:
						if (DateUtil.isCellDateFormatted(cell)) {
							Date date = cell.getDateCellValue();
							String cellString = new SimpleDateFormat("yyyy-MM-dd").format(date);
//                                System.out.println("The cell contains a date value: "
//                                        + cell.getDateCellValue());
							System.out.print(cellString + "\t");

						} else {
							cell.getNumericCellValue();
							System.out.print(+(long) cell.getNumericCellValue() + "\t");
						}
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case _NONE:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case BLANK:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case ERROR:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case FORMULA:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					}
				} // while
				System.out.println("");
			} // while
			file.close();

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return data;
	}

	public XSSFWorkbook makeExcel() throws FileNotFoundException, IOException {
		String path = "C:/Users/chox6/test";
		String fileName = "test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("아무거나1");
		XSSFRow curRow = sheet.createRow(0);
		XSSFCell cell = curRow.createCell(0);

		writeHeader("1", "1", sheet, curRow, cell);

		workbook.write(new FileOutputStream(path + "/" + fileName));
		workbook = new XSSFWorkbook(path + "/" + fileName);

		/*
		 * int rowindex = 0; int columnindex = 0; // 시트 수 (첫번째에만 존재하므로 0을 준다) // 만약 각
		 * 시트를 읽기위해서는 FOR문을 한번더 돌려준다 XSSFSheet sheet1 = workbook.getSheetAt(0); // 행의 수
		 * int rows = sheet.getPhysicalNumberOfRows(); for (rowindex = 0; rowindex <
		 * rows; rowindex++) { // 행을읽는다 XSSFRow row = sheet.getRow(rowindex); if (row !=
		 * null) { // 셀의 수 int cells = row.getPhysicalNumberOfCells(); for (columnindex
		 * = 0; columnindex <= cells; columnindex++) { // 셀값을 읽는다 XSSFCell cell1 =
		 * row.getCell(columnindex); String value = ""; // 셀이 빈값일경우를 위한 널체크 if (cell1 ==
		 * null) { continue; } else { // 타입별로 내용 읽기
		 * 
		 * // switch (cell1.getCellType()){ // case XSSFCell.CELL_TYPE_FORMULA: //
		 * value=cell.getCellFormula(); // break; // case XSSFCell.CELL_TYPE_NUMERIC: //
		 * value=cell.getNumericCellValue()+""; // break; // case
		 * XSSFCell.CELL_TYPE_STRING: // value=cell.getStringCellValue()+""; // break;
		 * // case XSSFCell.CELL_TYPE_BLANK: // value=cell.getBooleanCellValue()+""; //
		 * break; // case XSSFCell.CELL_TYPE_ERROR: //
		 * value=cell.getErrorCellValue()+""; // break; // } } //
		 * System.out.println(rowindex + "번 행 : " + columnindex + "번 열 값은: " +
		 * cell1.getStringCellValue());
		 * 
		 * 
		 * }
		 * 
		 * }
		 * 
		 * }
		 */
		return workbook;
	}
}