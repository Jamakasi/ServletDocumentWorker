/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package ru.jamakasi;

import java.io.IOException;
import java.util.Enumeration;
import java.util.HashMap;
import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 *
 * @author Admin
 */
public class main extends HttpServlet {
    protected WordprocessingMLPackage FindAndReplace(String docName,HashMap<String, String> mappings){
                ServletContext app = getServletContext();
                String inputfilepath = app.getRealPath("/files/test.docx");

                java.io.File tmp = null;
                WordprocessingMLPackage wordMLPackage = null;
                try{
                    tmp = new java.io.File(inputfilepath);
                    wordMLPackage = WordprocessingMLPackage.load(tmp);
                    VariablePrepare.prepare(wordMLPackage);
                    MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
                    documentPart.variableReplace(mappings);
                }catch(Exception e){

                }   
                return wordMLPackage;
    }
    protected void SendFile(HttpServletResponse response,WordprocessingMLPackage document, String outfileName ){
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-disposition", "attachment;filename="+outfileName+".docx");
        SaveToZipFile saver = new SaveToZipFile(document);
        try{
            saver.save( response.getOutputStream() );
        }catch(Exception e){
            
        }
    }
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
            throws ServletException, IOException 
    {
        /*
        * Перевариваем входные данные
        */
        String name,value,OutDocName="",WorkDocName="";
        HashMap<String, String> mappings = new HashMap<String, String>();
            Enumeration enumb = request.getParameterNames();
                for (; enumb.hasMoreElements(); ) {
                    // Get the name of the request parameter
                    name = (String)enumb.nextElement();
                    value = request.getParameter(name);
                    switch(name){
                        case "WorkDocName":{WorkDocName = value;break;}
                        case "OutDocName":{OutDocName=value; ;break;}
                        default :{mappings.put(name, value); break;}
                    }
                }
        /*
        * Теперь начинаем переваривать документы        
        */
        if(OutDocName.equals("")){
            OutDocName=WorkDocName;
        }
        SendFile(response, FindAndReplace(WorkDocName, mappings),OutDocName);

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
        return "Document generator";
    }// </editor-fold>

}
