/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;

import java.io.IOException;
import java.io.PrintWriter;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import com.google.gson.Gson;
import com.mysql.jdbc.Connection;
import com.mysql.jdbc.Statement;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Paths;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;
import javax.servlet.http.Part;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author turismo
 */
@WebServlet(name = "TestServlet",
        urlPatterns = {
            "/crear",
            "/subirExcel"})
@MultipartConfig
public class TestServlet extends HttpServlet {

    private final String UPLOADED_FILE_PATH = "d:\\";

    /**
     * Processes requests for both HTTP <code>GET</code> and <code>POST</code>
     * methods.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        try {
            if (request.getServletPath().equals("/crear")) {
                crear(request, response);
            } else if (request.getServletPath().equals("/subirExcel")) {
                subirExcel(request, response);
            } else {
                response.setContentType("text/html;charset=UTF-8");
                PrintWriter out = response.getWriter();
                try {
                    /* TODO output your page here. You may use following sample code. */
                    out.println("<!DOCTYPE html>");
                    out.println("<html>");
                    out.println("<head>");
                    out.println("<title>Servlet ControllerServlet</title>");
                    out.println("</head>");
                    out.println("<body>");
                    out.println("<h1>Servlet ControllerServlet at " + request.getContextPath() + "</h1>");
                    out.println("</body>");
                    out.println("</html>");
                } finally {
                    out.close();
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

    private void subirExcel(HttpServletRequest request, HttpServletResponse response) throws IOException, ServletException {
     try {
         
        Part filePart = request.getPart("file");
        String fileName = Paths.get(filePart.getSubmittedFileName()).getFileName().toString();
        InputStream fileContent = filePart.getInputStream();

        byte[] bytes = IOUtils.toByteArray(fileContent);

        // constructs upload file path
        fileName = UPLOADED_FILE_PATH + fileName;

        FileInputStream fileInputStream = writeFile(bytes, fileName);


        Workbook workBook = null;

   
            try {
                // Read XLSX document - Office 2007, 2010 format
                workBook = new XSSFWorkbook(fileInputStream);
            } catch (Exception e) {
                POIFSFileSystem fileSystem = new POIFSFileSystem(fileInputStream);
                // Read XLS document - Office 97 -2003 format
                workBook = new HSSFWorkbook(fileSystem);
            }

            int cantidadDeSolapas = workBook.getNumberOfSheets();
            for (int s = 0; s < cantidadDeSolapas; s++) {
                Sheet my_worksheet = workBook.getSheetAt(s);
                Iterator<Row> rowIterator = my_worksheet.iterator();

                String stringCell = "";
                Cell hssfCell;
                float f;
                int contador = 0;
                boolean isNumber = true;

                while (rowIterator.hasNext()) {
                    Row hssfRow = (Row) rowIterator.next();
                    Iterator<Cell> iterator = hssfRow.cellIterator();

                    while (iterator.hasNext()) {
                        hssfCell = (Cell) iterator.next();
                        stringCell = hssfCell.toString();

                    }

                }
            }
        } catch (Exception e) {
        }

        request.getRequestDispatcher("index.html").forward(request, response);
    }

    private void crear(HttpServletRequest request, HttpServletResponse response) throws IOException {

        try {
            Class.forName("com.mysql.jdbc.Driver").newInstance();
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException ex) {
            System.out.println("Error, no se ha podido cargar MySQL JDBC Driver");
        }

        String nombre = request.getParameter("nombre");
        String telefono = request.getParameter("telefono");

        try {

            String url = "jdbc:mysql://localhost:3306/mysql?zeroDateTimeBehavior=convertToNull";
            String username = "root";
            String password = "admin";

            Connection connection = (Connection) DriverManager.getConnection(url, username, password);

            Statement statement = (Statement) connection.createStatement();
            statement.executeUpdate("INSERT INTO Test VALUES ( \"" + nombre + "\", \"" + telefono + "\")");

            statement.close();
            connection.close();

        } catch (SQLException ex) {
            System.out.println(ex);
        }

        String json = new Gson().toJson("Ok");
        response.setContentType("application/json");
        response.getWriter().write(json);
    }

    private FileInputStream writeFile(byte[] content, String filename) throws IOException {

        File file = new File(filename);
        if (!file.exists()) {
            file.createNewFile();
        }
        FileUtils.writeByteArrayToFile(file, content);

        FileInputStream fileInputStream = new FileInputStream(file);

        return fileInputStream;
    }
}
