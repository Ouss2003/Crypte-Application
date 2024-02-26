package com.example.demo1;

import javafx.fxml.FXML;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class HelloController {

    @FXML
    private TextField pdfTextField; // Champ de texte pour le chemin du dossier PDF

    @FXML
    private TextField xslTextField; // Champ de texte pour le chemin du fichier Excel

    // Méthode appelée lorsque le bouton de cryptage est cliqué
    @FXML
    protected void onCryptButtonClick() {
        // Récupération des chemins des dossiers PDF et du fichier Excel à partir des champs de texte
        String pdfDirectoryPath = pdfTextField.getText();
        String excelFilePath = xslTextField.getText();

        // Vérification si les chemins ne sont pas vides
        if (pdfDirectoryPath.isEmpty() || excelFilePath.isEmpty()) {
            System.out.println("Veuillez sélectionner un dossier PDF et un fichier Excel.");
            return;
        }

        // Récupération de la liste des fichiers dans le dossier PDF spécifié
        File[] pdfFiles = new File(pdfDirectoryPath).listFiles();

        // Vérification si aucun fichier PDF n'a été trouvé dans le dossier spécifié
        if (pdfFiles == null || pdfFiles.length == 0) {
            System.out.println("Aucun fichier PDF trouvé dans le répertoire spécifié.");
            return;
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath))) {
            // Parcours de chaque fichier PDF dans le dossier
            for (File pdfFile : pdfFiles) {
                try (PDDocument document = PDDocument.load(pdfFile)) {
                    // Récupération du mot de passe associé au fichier PDF dans le fichier Excel
                    String password = getPasswordForPDF(workbook, pdfFile.getName());
                    if (password != null) {
                        // Cryptage du fichier PDF avec le mot de passe récupéré
                        document.protect(new StandardProtectionPolicy("", password, new AccessPermission()));
                        document.save(pdfFile);
                        System.out.println("Le fichier PDF '" + pdfFile.getName() + "' a été crypté avec succès.");
                    } else {
                        System.out.println("Aucun mot de passe trouvé pour le fichier PDF '" + pdfFile.getName() + "'.");
                    }
                }
            }
        } catch (IOException e) {
            System.out.println("Erreur lors du cryptage des fichiers PDF : " + e.getMessage());
        }
    }

    // Méthode pour récupérer le mot de passe associé au fichier PDF dans le fichier Excel
    private String getPasswordForPDF(XSSFWorkbook workbook, String pdfFileName) {
        // Parcours de chaque ligne dans la première feuille du classeur Excel
        for (Row row : workbook.getSheetAt(0)) {
            Cell cell = row.getCell(0);
            // Vérification si le nom du fichier PDF correspond au nom dans la première colonne
            if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equals(pdfFileName)) {
                Cell passwordCell = row.getCell(1);
                if (passwordCell != null && passwordCell.getCellType() == CellType.STRING) {
                    // Retourne le mot de passe si trouvé
                    return passwordCell.getStringCellValue();
                } else {
                    // Affiche un message si la cellule du mot de passe n'est pas une chaîne
                    System.out.println("La cellule du mot de passe pour le fichier PDF '" + pdfFileName + "' n'est pas une chaîne.");
                    return null;
                }
            }
        }
        // Retourne null si aucun mot de passe n'est trouvé pour le fichier PDF
        return null;
    }

    // Méthode pour sélectionner un dossier PDF
    @FXML
    protected void selectPdfDirectory() {
        // Création de la boîte de dialogue pour sélectionner un dossier
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Sélectionnez un Dossier PDF");

        // Affichage de la boîte de dialogue et récupération du dossier sélectionné par l'utilisateur
        File selectedDirectory = directoryChooser.showDialog(null);

        // Vérification si un dossier a été sélectionné
        if (selectedDirectory != null) {
            // Affichage du chemin absolu du dossier sélectionné
            String directoryPath = selectedDirectory.getAbsolutePath();
            System.out.println("Chemin du dossier PDF sélectionné : " + directoryPath);
            // Affichage du chemin dans le champ de texte de l'interface utilisateur
            pdfTextField.setText(directoryPath);
        }
    }

    // Méthode pour sélectionner un fichier Excel
    @FXML
    protected void selectXslFile() {
        // Création de la boîte de dialogue pour sélectionner un fichier
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Sélectionnez un fichier XSL");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Fichiers XSL (*.xlsx)", "*.xlsx"));

        // Affichage de la boîte de dialogue et récupération du fichier sélectionné par l'utilisateur
        File selectedFile = fileChooser.showOpenDialog(null);

        // Vérification si un fichier a été sélectionné et s'il existe
        if (selectedFile != null && selectedFile.exists()) {
            // Affichage du chemin absolu du fichier sélectionné
            String xslFilePath = selectedFile.getAbsolutePath();
            System.out.println("Chemin du fichier XSL sélectionné : " + xslFilePath);
            // Affichage du chemin dans le champ de texte de l'interface utilisateur
            xslTextField.setText(xslFilePath);
        } else {
            System.out.println("Aucun fichier XSL sélectionné.");
        }
    }
}
