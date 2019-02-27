package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
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

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

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

        Workbook workbook = new XSSFWorkbook();
        List<Facture> factures = factureService.findAllFacture();

        for (Facture facture:factures)
            {
                Sheet sheet = workbook.createSheet("Facture: "+facture.getId().toString());
                Row headerRow = sheet.createRow(0);
                Cell cellHeaderId = headerRow.createCell(0);
                cellHeaderId.setCellValue("début");

                Integer r=5;
                double total=0;

                for (LigneFacture ligne:facture.getLigneFactures()) {
                    Row row = sheet.createRow(r++);

                    Cell cell = row.createCell(1);
                    cell.setCellValue(ligne.getArticle().getLibelle());

                    cell = row.createCell(2);
                    double qte =ligne.getQuantite();
                    cell.setCellValue(qte);

                    cell = row.createCell(3);
                    double pu =ligne.getArticle().getPrix();
                    cell.setCellValue(pu);

                    cell = row.createCell(4);
                    pu =ligne.getArticle().getPrix();
                    cell.setCellValue(pu*qte);

                    total+=pu*qte;

                }

                r++;
                Row totalRow = sheet.createRow(0);
                Cell cell = totalRow.createCell(0);
                cell.setCellValue("Total");
                cell = totalRow.createCell(1);
                cell.setCellValue(total);


            }




        int i = 1;
        //for (Client client : allClients) {
        //   Row row = sheet.createRow(i);

        // i++;
        //}

        workbook.write(response.getOutputStream());
        workbook.close();

    }
}
