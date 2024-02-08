/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/JSP_Servlet/Servlet.java to edit this template
 */
package com.mycompany;

import java.io.IOException;
import java.io.PrintWriter;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.HttpConstraint;
import jakarta.servlet.annotation.ServletSecurity;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author manal
 */
@WebServlet(name = "ProcessXLSFileServlet", urlPatterns = {"/.xls"})
@ServletSecurity(
@HttpConstraint( rolesAllowed = {"xslacccess"}))
public class ProcessXLSFileServlet extends HttpServlet {

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
        response.setContentType("text/html;charset=UTF-8");
        try (PrintWriter out = response.getWriter()) {
          String fileName = request.getParameter("fileName");

          //  String filePath = "C:/Info5100_Labs/Program5/src/main/resources/" + fileName + ".xls";
            
            
            String filePath = getServletContext().getRealPath("/") + File.separator + "WEB-INF" + File.separator + "classes" + File.separator +  fileName + ".xls"; ;


            try (FileInputStream file = new FileInputStream(filePath);
                HSSFWorkbook workbook = new HSSFWorkbook(file)) {

                Sheet sheet = workbook.getSheetAt(0);

                Iterator<Row> rowIterator = sheet.iterator();
                out.println("<html><head><title>Read XLSX File</title></head><body><table border=\"1\">");

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    out.println("<tr>");

                    Iterator<Cell> cellItr = row.cellIterator();

                    while (cellItr.hasNext()) {
                        Cell cell = cellItr.next();
                        out.println("<td>");

                        switch (cell.getCellType()) {
                            case NUMERIC:
                                out.print(cell.getNumericCellValue());
                                break;
                            case STRING:
                                out.print(cell.getStringCellValue());
                                break;
                        }

                        out.println("</td>");
                    }

                    out.println("</tr>");
                }

                out.println("</table></body></html>");
            } catch (Exception e) {
                out.println("Error reading the file: " + e.getMessage());
            }
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

}
