package allbaro.cf.servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import com.oreilly.servlet.MultipartRequest;
import com.oreilly.servlet.multipart.DefaultFileRenamePolicy;
import com.oreilly.servlet.multipart.FileRenamePolicy;

import allbaro.cf.controller.AllbaroExcelController;
import allbaro.cf.controller.UploadUtil;

/**
 * Servlet implementation class AllbaroExcelServlet
 */

@MultipartConfig(maxFileSize = 1024 * 1024 * 5, maxRequestSize = 1024 * 1024 * 50)
@WebServlet("/AllbaroExcelServlet")
public class AllbaroExcelServlet extends HttpServlet {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
//		// 1. upload 폴더 생성이 안되어 있으면 생성
//		String saveDirectory = request.getServletContext().getRealPath("/excel");
//		System.out.println(saveDirectory);
//
//		File saveDir = new File(saveDirectory);
//		if (!saveDir.exists())
//			saveDir.mkdirs();
//
//		// 2. 최대크기 설정
//		int maxPostSize = 1024 * 1024 * 5; // 5MB 단위 byte
//
//		// 3. 인코딩 방식 설정
//		String encoding = "UTF-8";
//
//		
//		// 4. 파일정책, 파일이름 충동시 덮어씌어짐으로 파일이름 뒤에 인덱스를 붙인다.
//		// a.txt
//		// a1.txt 와 같은 형식으로 저장된다.
//		FileRenamePolicy policy = new DefaultFileRenamePolicy();
//		MultipartRequest mrequest = new MultipartRequest(request // MultipartRequest를 만들기 위한 request
//				, saveDirectory // 저장 위치
//				, maxPostSize // 최대크기
//				, encoding // 인코딩 타입
//				, policy); // 파일 정책
//
//		String name = mrequest.getParameter("name");
//		File uploadFile = mrequest.getFile("excel");
//		// input type="file" 태그의 name속성값을 이용해 파일객체를 생성
//		long uploadFile_length = uploadFile.length();
//		String originalFileName = mrequest.getOriginalFileName("excel"); // 기존 이름
//		String filesystemName = mrequest.getFilesystemName("excel"); // 기존

		AllbaroExcelController aec = new AllbaroExcelController();

//		aec.getData(new FileInputStream("C:/Users/chox6/OneDrive/바탕 화면/allbaro/Excel (3).xlsx"));
		

//		aec.makeExcel();
		
		// https://dev-gorany.tistory.com/289
		//UploadUtil uploadUtil = UploadUtil.create(request.getServletContext());
		
		//List<Part> parts = (List<Part>) request.getParts();
//		Part part = request.getPart(getServletName());
		
//		for(Part part : parts) {
//			if(!part.getName().equals("excel")) continue;// excel로 들어온 part가 아니면 스킵
//			if(part.getSubmittedFileName().equals(""))continue; // 업로드된 파일 이름이 없으면 스킵
//			
//			String fileName = part.getSubmittedFileName();
//			
//			uploadUtil.saveFiles(part, "");
//			
//		}
		
		ServletContext context = request.getSession().getServletContext();
		String realPath = context.getRealPath("/excel");
		int maxSize = 1024*1024*30;
		
		MultipartRequest multi = new MultipartRequest(request, realPath, maxSize, "utf-8",new DefaultFileRenamePolicy());
		String fileName = multi.getFilesystemName("excel");
		String originalName = multi.getOriginalFileName("excel");
		String type = multi.getContentType("excel");
		File file = multi.getFile("excel");
		
		System.out.println(file.getName());
		System.out.println(file.getAbsolutePath());
		System.out.println(file.length());
		
		aec.outputExcel(aec.getData(file,realPath));


		}
}
