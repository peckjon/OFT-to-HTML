import com.aspose.email.License;
import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
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

        // Directory containing OFT files
        File inputDir = new File("input");

        // Output directory for HTML files
        File outputDir = new File("output");
        if (!outputDir.exists()) {
            // Create output directory if it doesn't exist
            outputDir.mkdirs();
        }

        // List all OFT files in the input directory
        FilenameFilter oftFilter = (dir, name) -> name.toLowerCase().endsWith(".oft");
        File[] oftFiles = inputDir.listFiles(oftFilter);

        // Convert OFT files to HTML
        if (oftFiles != null) {
            for (File oftFile : oftFiles) {
                try {
                    // Convert OFT file to HTML
                    MailMessage message = MailMessage.load(oftFile.getAbsolutePath());
                    String htmlOutputPath = outputDir.getAbsolutePath() + File.separator + oftFile.getName().replace(".oft", ".html");
                    message.save(htmlOutputPath, SaveOptions.getDefaultHtml());
                    System.out.println("Converted: " + oftFile.getName() + " to HTML at " + htmlOutputPath);
                } catch (Exception e) {
                    System.err.println("An error occurred during conversion of " + oftFile.getName() + ": " + e.getMessage());
                }
            }
        } else {
            System.err.println("No OFT files found in the input directory.");
        }
    }
}