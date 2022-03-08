package allbaro.cf.controller;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;

import allbaro.cf.dto.CF_dto;

public class AllbaroExcelController {
	
	//추후 수정
	public void writeDoc() {
		String tmpDate = "10월";
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//tmpDate로 빈 시트생성
		XSSFSheet sheet = workbook.createSheet(tmpDate);
		//data담을 곳
		 Map<String, CF_dto[]> data = makeDummy();
		
		int rownum = 0;
        for (String key : data)
        {
            Row row = sheet.createRow(rownum++);
            CF_dto [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
	}
	//make dummy data
		public TreeMap<String, CF_dto[]> makeDummy(){
			 Map<String, CF_dto[]> data = new TreeMap<String, CF_dto[]>();
			//dummy data
			
			data.put(new CF_dto("2021-10-12","종류1","123123","112233445","A company",50,0,"11223344"));
			data.add(new CF_dto("2021-10-12","종류1","123123","112233445","A company",0,50,"11223344"));
			data.add(new CF_dto("2021-10-14","종류2","321321","123456789","A company",2000,0,"44332211"));
			data.add(new CF_dto("2021-10-14","종류2","321321","123456789","B company",0,2000,"44332211"));
			data.add(new CF_dto("2021-10-15","종류1","123123","112233445","A company",30,0,"12312312"));
			data.add(new CF_dto("2021-10-15","종류1","123123","112233445","A company",0,30,"12312312"));
			data.add(new CF_dto("2021-10-19","종류1","123123","112233445","A company",90,0,"32132132"));
			data.add(new CF_dto("2021-10-19","종류1","123123","112233445","A company",0,90,"32132132"));
			data.add(new CF_dto("2021-10-19","종류1","123123","987654321","D company",900,0,"32132112"));
			data.add(new CF_dto("2021-10-19","종류1","123123","987654321","D company",0,900,"32132112"));
			data.add(new CF_dto("2021-10-20","종류2","321321","123456789","B company",8000,0,"32132155"));
			data.add(new CF_dto("2021-10-20","종류2","321321","123456789","B company",0,8000,"32132155"));
			data.add(new CF_dto("2021-10-25","종류1","123123","112233445","A company",100,0,"32132166"));
			data.add(new CF_dto("2021-10-26","종류3","112233","111111111","C company",200,0,"32132177"));
			data.add(new CF_dto("2021-10-26","종류3","112233","111111111","C company",0,200,"32132177"));
			data.add(new CF_dto("2021-10-26","종류1","123123","112233445","A company",0,100,"32132166"));
			data.add(new CF_dto("2021-10-26","종류1","123123","987654321","D company",500,0,"32132188"));
			data.add(new CF_dto("2021-10-26","종류1","123123","987654321","D company",0,500,"32132188"));
			data.add(new CF_dto("2021-10-26","종류4","132132","111111111","C company",600,0,"32132199"));
			data.add(new CF_dto("2021-10-26","종류4","132132","111111111","C company",0,600,"32132199"));
			data.add(new CF_dto("2021-10-26","종류2","321321","123456789","B company",3000,0,"32132100"));
			data.add(new CF_dto("2021-10-27","종류2","321321","123456789","B company",0,3000,"32132100"));
			data.add(new CF_dto("2021-10-29","종류1","123123","112233445","A company",0,0,"11223344"));
			data.add(new CF_dto("2021-10-29","종류1","123123","112233445","A company",0,0,"11223344"));
			
			return data;		
		}
	
	
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
	
	
	public void writeHeader() {
		
	}//writeHeader
	public void makeExcel() throws FileNotFoundException, IOException {
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
	}
}
