package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

// Die Klasse FileVisitor erweitert SimpleFileVisitor, um durch ein Dateisystem zu navigieren
public class FileVisitor extends SimpleFileVisitor<Path> {
    // Der Pfad des Verzeichnisses, das durchsucht werden soll
    static String directoryPath = "C:\\Users\\Student\\Desktop\\Neuer Ordner (3)";
    // Der Pfad des Eingabeverzeichnisses
    String inFilePath = "C:\\Users\\Student\\Desktop\\Neuer Ordner (3)\\";

    // Listen zum Speichern von Dateipfaden
    List<String> filePaths = new ArrayList<>();
    List<String> newFilePaths = new ArrayList<>();

    public static void main(String[] args) throws IOException {
        // Erstellt eine Instanz von FileVisitor
        FileVisitor visitor = new FileVisitor();

        // Durchläuft das Dateisystem ab dem angegebenen Pfad und besucht jede Datei
        Files.walkFileTree(Paths.get(FileVisitor.directoryPath), visitor);

        // Ruft die Ausgabe- und Änderungsmethoden auf
        visitor.ausgabe();
        visitor.changeFile();
    }

    // Überschreibt die visitFile-Methode, um gefundene Dateien zu verarbeiten
    @Override
    public FileVisitResult visitFile(Path file, BasicFileAttributes att) throws IOException {
        // Fügt den Dateinamen der Liste filePaths hinzu
        filePaths.add(file.getFileName().toString());
        return FileVisitResult.CONTINUE; // Setzt die Dateisuche fort
    }

    // Fügt den vollständigen Pfad zu jeder gefundenen Datei zur Liste newFilePaths hinzu
    void ausgabe() {
        for (String s : filePaths) {
            newFilePaths.add(inFilePath + s);
        }

        // Kommentierter Code: Der Code unten kann verwendet werden, um die neuen Dateipfade auszugeben
        // for (String s1 : newFilePaths)
        //     System.out.println();
    }

    /*
    // Methode, um die Anzahl der Dateien in einem Verzeichnis zu zählen
    public int dirCount(){
        File directory = new File("C:\\Users\\Student\\Desktop\\Neuer Ordner (3)"); // Das Verzeichnis, das betrachtet wird
        String[] list = directory.list(); // Liste der Dateien im Verzeichnis
        int count = list.length; // Anzahl der gefundenen Dateien
        System.out.println(count);
        return count;
    }
    */

    // Diese Methode nimmt einen Dateipfad als Parameter und filtert den Inhalt der Datei
    void ausgabe2(String s) {
        // Liste der Wörter, die gefiltert werden sollen
        List<String> wordsToFilter = Arrays.asList("Uwe Wagner");
        try {
            // Öffnet die Datei für den Lesezugriff
            FileInputStream fis = new FileInputStream(s);
            XWPFDocument doc = new XWPFDocument(fis);

            // Ruft alle Tabellen im Dokument ab
            List<XWPFTable> tables = doc.getTables();

            // Durchläuft jede Tabelle im Dokument
            for (XWPFTable table : tables) {
                System.out.println("Tabelle gefunden:");

                // Durchläuft jede Zeile in der Tabelle
                for (XWPFTableRow row : table.getRows()) {

                    // Durchläuft jede Zelle in der Zeile
                    for (XWPFTableCell cell : row.getTableCells()) {
                        // Textinhalt der Zelle abrufen
                        String text = cell.getText();
                        String originalText = text; // Speichert den ursprünglichen Text

                        // Überprüft, ob der Text Wörter aus der Filterliste enthält und ersetzt diese
                        for (String word : wordsToFilter) {
                            if (text.contains(word)) {
                                text = text.replace(word, "Christopher Briesemann");
                            }
                        }
                        // Setzt den Text nur neu, wenn er geändert wurde
                        if (!text.equals(originalText)) {
                            cell.removeParagraph(0); // Entfernt den alten Paragraphen
                            cell.setText(text); // Setzt den neuen Text
                        }

                        // Durchläuft alle Absätze im Dokument
                        for (XWPFParagraph paragraph : doc.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                List<XWPFPicture> pictures = run.getEmbeddedPictures();

                                // Entfernt eingebettete Bilder, falls vorhanden
                                if (!pictures.isEmpty()) {
                                    int drawingCount = run.getCTR().sizeOfDrawingArray();
                                    if (drawingCount > 0) {
                                        run.getCTR().removeDrawing(0); // Entfernt das erste Drawing-Element
                                    }
                                }
                            }
                        }
                        System.out.println("Zelleninhalt: " + cell.getText());
                    }
                }
            }

            // Speichert das geänderte Dokument
            FileOutputStream fos = new FileOutputStream(s);
            doc.write(fos);

            // Schließt die Ressourcen
            fos.close();
            doc.close();
            fis.close();

            System.out.println("Das Dokument wurde erfolgreich gefiltert und gespeichert.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Durchläuft die Liste der neuen Dateipfade und wendet die ausgabe2-Methode darauf an
    void changeFile() {
        for (String s : newFilePaths) {
            ausgabe2(s);
        }
    }
}
