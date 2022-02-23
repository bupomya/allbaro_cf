package allbaro.cf.controller;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AllbaroExcelController {

	// 경로에 있는 엑셀 파일 읽어오기
	public void readExcel(String dir) {
		String fileName = "testExcel.xlsx";

	}// readExcel

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
