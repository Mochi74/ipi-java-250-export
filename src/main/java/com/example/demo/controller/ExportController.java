package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;


    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");
        LocalDate now = LocalDate.now();
        for (Client client : allClients) {
            writer.println(
                    client.getId() + ";"
                            + client.getNom() + ";"
                            + client.getPrenom() + ";"
                            + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                            + (now.getYear() - client.getDateNaissance().getYear())
            );
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Clients");


        Row headerRow = sheet.createRow(0);

        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        Cell cellHeaderPrenom = headerRow.createCell(1);
        cellHeaderPrenom.setCellValue("Prénom");

        Cell cellHeaderNom = headerRow.createCell(2);
        cellHeaderNom.setCellValue("Nom");

        Cell cellHeaderDateNaissance = headerRow.createCell(3);
        cellHeaderDateNaissance.setCellValue("Date de naissance");

        int i = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(i);

            Cell cellId = row.createCell(0);
            cellId.setCellValue(client.getId());

            Cell cellPrenom = row.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = row.createCell(2);
            cellNom.setCellValue(client.getNom());

            Cell cellDateNaissance = row.createCell(3);
            Date dateNaissance = Date.from(client.getDateNaissance().atStartOfDay(ZoneId.systemDefault()).toInstant());
            cellDateNaissance.setCellValue(dateNaissance);

            CellStyle cellStyleDate = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
            cellDateNaissance.setCellStyle(cellStyleDate);

            i++;
        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        XSSFWorkbook workbook = new XSSFWorkbook();
        List<Client> allClients = clientService.findAllClients();

      for (Client client: allClients) {
          XSSFSheet clientSheet = workbook.createSheet("Client " + client.getNom().toString());

          List<Facture> factures = factureService.findAllFactureByClient(client);

          Row clientRow = clientSheet.createRow(3);
          Cell cellClientId = clientRow.createCell(1);
          cellClientId.setCellValue("Facture pour :");
          cellClientId = clientRow.createCell(3);
          cellClientId.setCellValue(client.getPrenom() + " " + client.getNom());

          //Create a new font and alter it.
          XSSFFont font = workbook.createFont();
          font.setBold(true);
          font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());

          //Set font into style
          XSSFCellStyle styleBold = workbook.createCellStyle();
          styleBold.setFont(font);
          styleBold.setBorderBottom(BorderStyle.THICK);
          styleBold.setBorderTop(BorderStyle.THICK);
          styleBold.setBorderRight(BorderStyle.THICK);
          styleBold.setBorderLeft(BorderStyle.THICK);


          for (Facture facture : factures) {

              XSSFSheet sheet = workbook.createSheet("Facture " + facture.getId().toString());
              Integer r = 3;

              // edit client
              // add an empty line
              r++;
              // build entete
              Row headerRow = sheet.createRow(r++);
              Cell cellHeaderId = headerRow.createCell(1);
              cellHeaderId.setCellValue("Libelle");
              cellHeaderId = headerRow.createCell(2);
              cellHeaderId.setCellValue("Quantité");
              cellHeaderId = headerRow.createCell(3);
              cellHeaderId.setCellValue("Prix Unitaire");
              cellHeaderId = headerRow.createCell(4);
              cellHeaderId.setCellValue("Prix");

              //build facture lines
              double total = 0;

              for (LigneFacture ligne : facture.getLigneFactures()) {
                  Row row = sheet.createRow(r++);

                  Cell cell = row.createCell(1);
                  cell.setCellValue(ligne.getArticle().getLibelle());

                  cell = row.createCell(2);
                  double qte = ligne.getQuantite();
                  cell.setCellValue(qte);

                  cell = row.createCell(3);
                  double pu = ligne.getArticle().getPrix();
                  cell.setCellValue(pu);

                  cell = row.createCell(4);
                  pu = ligne.getArticle().getPrix();
                  cell.setCellValue(pu * qte);

                  total += pu * qte;

              }

              r++;
              //add total
              Row totalRow = sheet.createRow(r);
              sheet.addMergedRegion(new CellRangeAddress(r, r, 1, 3));
              Cell cell = totalRow.createCell(1);
              cell.setCellValue("Total");
              cell = totalRow.createCell(4);
              cell.setCellValue(total);
              cell.setCellStyle(styleBold);
          }
      }
        workbook.write(response.getOutputStream());
        workbook.close();

    }



    @GetMapping("/factures/{id}/pdf")
    public void facturesPdf(HttpServletRequest request, HttpServletResponse response,@PathVariable String id) throws DocumentException, IOException {
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures"+id+".pdf\"");

        Document document = new Document();
        PdfWriter.getInstance(document, response.getOutputStream());
        document.open();
        document.add(new Paragraph("Hello World! "+id));


        List<Facture> factures = factureService.findAllFacture();

        for (Facture facture:factures) {
            document.add(new Paragraph(facture.getClient().getPrenom() + " " + facture.getClient().getNom()));
            double qte;
            double pu;
            double total=0;
            document.add(new Paragraph("Libellé Quantité Prix Unitaire Prix"));
            for (LigneFacture ligne:facture.getLigneFactures()) {
                ligne.getArticle().getLibelle();
                qte =ligne.getQuantite();
                pu =ligne.getArticle().getPrix();
                document.add(new Paragraph(ligne.getArticle().getLibelle()+ qte + pu + qte*pu));
                total+=pu*qte;
            }

        }
        document.close();
    }
}

