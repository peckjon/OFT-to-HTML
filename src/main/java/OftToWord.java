import com.aspose.email.License;
import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import java.io.File;
import java.io.FilenameFilter;

public class OftToWord {
    public static void main(String[] args) {

        // Set the license for Aspose.Email
        try {
            new License().setLicense("Aspose.EmailforJava.lic");
        } catch (Exception e) {
            System.err.println("Failed to load Aspose.Email license: " + e.getMessage());
        }

        // Set the license for Aspose.Words
        try {
            new License().setLicense("Aspose.WordsforJava.lic");
        } catch (Exception e) {
            System.err.println("Failed to load Aspose.Email license: " + e.getMessage());
        }

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
                    String htmlOutputPath = outputDir.getAbsolutePath() + File.separator + oftFile.getName().replace(".oft", ".html");

                    // Save OFT as HTML
                    message.save(htmlOutputPath, SaveOptions.getDefaultHtml());
                    
                    // Load the HTML file with an instance of Document
                    Document document = new Document(htmlOutputPath);
                    
                    // Construct output DOCX file path
                    String outputFilePath = outputDir.getAbsolutePath() + File.separator + oftFile.getName().replace(".oft", ".docx");
                    
                    // Save the document in DOCX format
                    document.save(outputFilePath, SaveFormat.DOCX);
                    System.out.println("Converted: " + oftFile.getName() + " to " + outputFilePath);

                } catch (Exception e) {
                    System.err.println("An error occurred during conversion of " + oftFile.getName() + ": " + e.getMessage());
                }
            }
        } else {
            System.err.println("No OFT files found in the input directory.");
        }
    }
}