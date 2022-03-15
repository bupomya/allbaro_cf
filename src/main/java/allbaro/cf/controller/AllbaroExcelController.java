package allbaro.cf.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

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

	public void outputExcel(Map<Integer, CF_dto> data) {
		
	}
	//
	public ArrayList<CF_dto> getInputData(ArrayList<CF_dto> all,int flag){
		//flag로 수집날짜,재활용날짜 data구분 1:수집 0:재활용
		ArrayList<CF_dto> data = new ArrayList<CF_dto>();
		
		switch(flag) {
		case 1://수집data
			if(all.get(1).getInAmount()!=0) {
				data.add(all.get(1));
			}
			break;
		case 2://재활용data
			if(all.get(1).getOutAmount()!=0) {
				data.add(all.get(1));
			}
			break;
		}
		return data;
	}
	
	// make dummy data
	public ArrayList<CF_dto> makeDummy() {
		ArrayList<CF_dto> data = new ArrayList<CF_dto>();
		// dummy data
		data.add(new CF_dto("2021-10-12", "종류1", "123123", "112233445", "A company", 50, 0, "11223344"));
		data.add(new CF_dto("2021-10-12", "종류1", "123123", "112233445", "A company", 0, 50, "11223344"));
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
		data.add(new CF_dto("2021-10-29", "종류1", "123123", "112233445", "A company", 0, 0, "11223344"));
		data.add(new CF_dto("2021-10-29", "종류1", "123123", "112233445", "A company", 0, 0, "11223344"));

		return data;
	}
	

	// 경로에 있는 엑셀 파일 읽어오기

	public void writeHeader(String year, String month,XSSFSheet xssfSheet, XSSFRow xssfRow, XSSFCell xssfCell) throws FileNotFoundException, IOException {

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
	
	
	
	public void writeData(FileInputStream file) {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			int rowIndex = 0;
			int cellIndex = 0;
			
			XSSFSheet sheet = workbook.getSheetAt(0); 
			int rows = sheet.getPhysicalNumberOfRows();
			
			//여기서 부터
/*	          for(rowindex=0;rowindex<rows;rowindex++){
	                //행을읽는다
	                XSSFRow row=sheet.getRow(rowindex);
	                if(row !=null){
	                    //셀의 수
	                    int cells=row.getPhysicalNumberOfCells();
	                    for(columnindex=0; columnindex<=cells; columnindex++){
	                        //셀값을 읽는다
	                        XSSFCell cell=row.getCell(columnindex);
	                        String value="";
	                        //셀이 빈값일경우를 위한 널체크
	                        if(cell==null){
	                            continue;
	                        }else{
	                            //타입별로 내용 읽기
	                            switch (cell.getCellType()){
	                            case XSSFCell.CELL_TYPE_FORMULA:
	                                value=cell.getCellFormula();
	                                break;
	                            case XSSFCell.CELL_TYPE_NUMERIC:
	                                value=cell.getNumericCellValue()+"";
	                                break;
	                            case XSSFCell.CELL_TYPE_STRING:
	                                value=cell.getStringCellValue()+"";
	                                break;
	                            case XSSFCell.CELL_TYPE_BLANK:
	                                value=cell.getBooleanCellValue()+"";
	                                break;
	                            case XSSFCell.CELL_TYPE_ERROR:
	                                value=cell.getErrorCellValue()+"";
	                                break;
	                            }
	                        }
	                        System.out.println(rowindex+"번 행 : "+columnindex+"번 열 값은: "+value);
	                    }
	 
	                }
	            }
			*/
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}


	public XSSFWorkbook makeExcel() throws FileNotFoundException, IOException {
		String path = "C:/Users/chox6/test";
		String fileName = "test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("아무거나1");
		XSSFRow curRow = sheet.createRow(0);
		XSSFCell cell = curRow.createCell(0);
		writeHeader("1","1",sheet, curRow,cell);

		workbook.write(new FileOutputStream(path + "/" + fileName));
		workbook = new XSSFWorkbook(path + "/" + fileName);

		/*
		int rowindex = 0;
		int columnindex = 0;
		// 시트 수 (첫번째에만 존재하므로 0을 준다)
		// 만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		// 행의 수
		int rows = sheet.getPhysicalNumberOfRows();
		for (rowindex = 0; rowindex < rows; rowindex++) {
			// 행을읽는다
			XSSFRow row = sheet.getRow(rowindex);
			if (row != null) {
				// 셀의 수
				int cells = row.getPhysicalNumberOfCells();
				for (columnindex = 0; columnindex <= cells; columnindex++) {
					// 셀값을 읽는다
					XSSFCell cell1 = row.getCell(columnindex);
					String value = "";
					// 셀이 빈값일경우를 위한 널체크
					if (cell1 == null) {
						continue;
					} else {
						// 타입별로 내용 읽기

//                        switch (cell1.getCellType()){
//                        case XSSFCell.CELL_TYPE_FORMULA:
//                            value=cell.getCellFormula();
//                            break;
//                        case XSSFCell.CELL_TYPE_NUMERIC:
//                            value=cell.getNumericCellValue()+"";
//                            break;
//                        case XSSFCell.CELL_TYPE_STRING:
//                            value=cell.getStringCellValue()+"";
//         
						break;
//                        case XSSFCell.CELL_TYPE_BLANK:
//                            value=cell.getBooleanCellValue()+"";
//                            break;
//                        case XSSFCell.CELL_TYPE_ERROR:
//                            value=cell.getErrorCellValue()+"";
//                            break;
//                        }
					}
//					System.out.println(rowindex + "번 행 : " + columnindex + "번 열 값은: " + cell1.getStringCellValue());
 * 
 
				}

			}
			
		}*/
		return workbook;
	}
}