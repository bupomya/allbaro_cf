package allbaro.cf.controller;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import allbaro.cf.dto.CF_dto;

public class AllbaroExcelController {
	CF_dto tmp = new CF_dto("2021-10-12","2021-10-12","종류1","123123","112233445","A company",50,50,"11223344");
	// 경로에 있는 엑셀 파일 읽어오기
	public JSONArray readExcel(MultipartFile excelFile) {
        JSONArray jsonArray = new JSONArray();
        
        try{
            XSSFWorkbook workbook = new XSSFWorkbook(excelFile.getInputStream());
            XSSFSheet sheet;
            XSSFRow curRow;
            XSSFCell curCell;
            
            /////////첫번째 시트 읽기
            sheet = workbook.getSheetAt(0);
            
            /////////행의 개수 
            int rownum = sheet.getLastRowNum();
            
            /////////첫 행은 컬럼명이므로 두번째 행인 index1부터 읽기
            for(int rowIndex=1;rowIndex<=rownum;rowIndex++){
            
                /////////현재 index의 행 읽기
                curRow = sheet.getRow(rowIndex);
                
                /////////현재 행의 cell 개수
                int cellnum = curRow.getLastCellNum();
                
                /////////엑셀 데이터를 넣을 json object
                JSONObject data = new JSONObject();
                
                for(int cellIndex =0; cellIndex<cellnum;cellIndex++){
                    System.out.println(rowIndex+"행 "+cellIndex+"열의 값 : " + curRow.getCell(cellIndex));
                    switch(cellIndex){
                        case 0 : {    /////////첫번째 열 값
                            data.put("번호",curRow.getCell(cellIndex).getStringCellValue());
                        };break;
                        case 1 : {    /////////두번째 열 값
                            data.put("이름",curRow.getCell(cellIndex).getStringCellValue());
                        };break;
                        case 2 : {    /////////세번째 열 값
                            data.put("주소",curRow.getCell(cellIndex).getStringCellValue());
                        }break;
                    }
                }

                jsonArray.add(data);
            }
            return jsonArray;
        }catch(Exception e){
            System.out.println("ERROR : "  + e);
            return jsonArray;
        }
    }
	
	
	public void writeHeader(String year, String month) throws FileNotFoundException, IOException {
		XSSFWorkbook xssfWb = makeExcel();
		XSSFSheet xssfSheet = xssfWb.getSheet("아무거나1");
		XSSFRow xssfRow = xssfSheet.getRow(0);
		
		XSSFCell xssfCell = xssfRow.getCell(0);//

		int rowNo = 0;
		
		
		xssfRow = xssfSheet.createRow(rowNo++); // 행 객체 추가
		xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
		xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8)); // 첫행, 마지막행,첫열,마지막열
		// 타이틀 생성
		xssfCell.setCellValue("[별지 제45호서식] <개정 2008.8.4>");
		
		//(앞쪽)
		xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 13));
		
		
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 13));
		xssfCell = xssfRow.getCell(1);
		xssfCell.setCellValue("폐기물 수탁 재활용 관리대장");
		
		xssfSheet.addMergedRegion(new CellRangeAddress(2, 2, 2, 13));
		xssfRow = xssfSheet.createRow(rowNo++);

		xssfCell.setCellValue("(재활용대상 폐기물의 종류:)(단위: 톤)");
		
		
		//year
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short) 0);
		xssfCell.setCellValue(year);
		
		xssfCell = xssfRow.createCell((short) 1);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 1, 3, 2));
		xssfCell.setCellValue("수집ㆍ운반내용");
		
		xssfCell = xssfRow.createCell((short) 3);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 5));
		xssfCell.setCellValue("폐기물재활용내용");
		
		xssfCell = xssfRow.createCell((short)6);
		xssfCell.setCellValue("수탁"); // "수탁" 폐기물 보관량
		
		xssfCell = xssfRow.createCell((short)7);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 7, 3, 12));
		xssfCell.setCellValue("재활용제품의 공급 및 보관내용");
		
		xssfCell = xssfRow.createCell((short)13);
		xssfSheet.addMergedRegion(new CellRangeAddress(3, 13, 5, 13));
		xssfCell.setCellValue("결재");
		
		
		
		//month
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short)0);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,0,5,0));
		xssfCell.setCellValue(month);
		
		
		xssfCell = xssfRow.createCell((short)1);
		xssfCell.setCellValue("수집");
		
		xssfCell = xssfRow.createCell((short)2);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,2,5,2));
		xssfCell.setCellValue("수집량");
		
		xssfCell = xssfRow.createCell((short)3);
		xssfCell.setCellValue("폐기물");//"폐기물" 재활용량
		
		xssfCell = xssfRow.createCell((short)4);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,4,4,5));
		xssfCell.setCellValue("재활용제품");
		
		xssfCell = xssfRow.createCell((short)6);
		xssfCell.setCellValue("폐기물");// 수탁 "폐기물" 보관량
		
		xssfCell = xssfRow.createCell((short)7);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,7,4,10));
		xssfCell.setCellValue("공급처");
		
		xssfCell = xssfRow.createCell((short)11);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,11,5,11));
		xssfCell.setCellValue("공급량");
		
		xssfCell = xssfRow.createCell((short)12);
		xssfSheet.addMergedRegion(new CellRangeAddress(4,12,5,12));
		xssfCell.setCellValue("보관량");
		
		
		xssfRow = xssfSheet.createRow(rowNo++);
		xssfCell = xssfRow.createCell((short)1);
		xssfCell.setCellValue("업소명");
		
		xssfCell = xssfRow.createCell((short)3);
		xssfCell.setCellValue("재활용량");// 폐기물 "재활용량"
		
		xssfCell = xssfRow.createCell((short)4);
		xssfCell.setCellValue("종류");
		
		xssfCell = xssfRow.createCell((short)5);
		xssfCell.setCellValue("생산량");
		
		xssfCell = xssfRow.createCell((short)6);
		xssfCell.setCellValue("보관량");// 수탁 폐기물 "보관량"
		
		xssfCell = xssfRow.createCell((short)7);
		xssfCell.setCellValue("상호(대표자)");
		
		xssfCell = xssfRow.createCell((short)8);
		xssfSheet.addMergedRegion(new CellRangeAddress(5,8,5,9));
		xssfCell.setCellValue("소재지");
		
		xssfCell = xssfRow.createCell((short)10);
		xssfCell.setCellValue("사용용도");
		
	}// writeHeader
	
	public XSSFWorkbook makeExcel() throws FileNotFoundException, IOException {
		String path = "C:/Users/chox6/test";
		String fileName = "test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		XSSFSheet sheet = workbook.createSheet("아무거나1");
		XSSFRow curRow = sheet.createRow(0);
		XSSFCell cell = curRow.createCell(0);
		cell.setCellValue("데이터");

		workbook.write(new FileOutputStream(path+"/"+fileName));
		workbook = new XSSFWorkbook(path+"/"+fileName);

		int rowindex=0;
        int columnindex=0;
        //시트 수 (첫번째에만 존재하므로 0을 준다)
        //만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
        XSSFSheet sheet1=workbook.getSheetAt(0);
        //행의 수
        int rows=sheet.getPhysicalNumberOfRows();
        for(rowindex=0;rowindex<rows;rowindex++){
            //행을읽는다
            XSSFRow row=sheet.getRow(rowindex);
            if(row !=null){
                //셀의 수
                int cells=row.getPhysicalNumberOfCells();
                for(columnindex=0; columnindex<=cells; columnindex++){
                    //셀값을 읽는다
                    XSSFCell cell1=row.getCell(columnindex);
                    String value="";
                    //셀이 빈값일경우를 위한 널체크
                    if(cell1==null){
                        continue;
                    }else{
                        //타입별로 내용 읽기
                    	
//                        switch (cell1.getCellType()){
//                        case XSSFCell.CELL_TYPE_FORMULA:
//                            value=cell.getCellFormula();
//                            break;
//                        case XSSFCell.CELL_TYPE_NUMERIC:
//                            value=cell.getNumericCellValue()+"";
//                            break;
//                        case XSSFCell.CELL_TYPE_STRING:
//                            value=cell.getStringCellValue()+"";
//                            break;
//                        case XSSFCell.CELL_TYPE_BLANK:
//                            value=cell.getBooleanCellValue()+"";
//                            break;
//                        case XSSFCell.CELL_TYPE_ERROR:
//                            value=cell.getErrorCellValue()+"";
//                            break;
//                        }
                    }
                    System.out.println(rowindex+"번 행 : "+columnindex+"번 열 값은: "+cell1.getStringCellValue());
                }

            }
        }
        return workbook;
	}
}