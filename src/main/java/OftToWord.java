import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import java.io.File;
import java.io.FilenameFilter;

public class OftToWord {
    public static void main(String[] args) {
        // Directory containing OFT files
        File inputDir = new File("input");
        // Output directory for DOCX files
        File outputDir = new File("output");
        if (!outputDir.exists()) {
            outputDir.mkdirs(); // Create output directory if it doesn't exist
        }

        // Filter to identify OFT files
        FilenameFilter oftFilter = (dir, name) -> name.toLowerCase().endsWith(".oft");

        // List all OFT files in the input directory
        File[] oftFiles = inputDir.listFiles(oftFilter);

        if (oftFiles != null) {
            for (File oftFile : oftFiles) {
                try {
                    // Load the OFT file
                    MailMessage message = MailMessage.load(oftFile.getAbsolutePath());
                    
                    // Temporary HTML output file path
                    String htmlOutputPath = "HtmlOutput.html";
                    // Save OFT as HTML
                    message.save(htmlOutputPath, SaveOptions.getDefaultHtml());
                    
                    // Load the HTML file with an instance of Document
                    Document document = new Document(htmlOutputPath);
                    
                    // Construct output DOCX file path
                    String outputFilePath = outputDir.getAbsolutePath() + File.separator + oftFile.getName().replace(".oft", ".docx");
                    // Save the document in DOCX format
                    document.save(outputFilePath, SaveFormat.DOCX);
                    System.out.println("Converted: " + oftFile.getName() + " to " + outputFilePath);
                    //delete the temporary HTML file
                    new File(htmlOutputPath).delete(); // Delete the temporary HTML file
                } catch (Exception e) {
                    System.err.println("An error occurred during conversion of " + oftFile.getName() + ": " + e.getMessage());
                }
            }
        } else {
            System.err.println("No OFT files found in the input directory.");
        }
    }
}