import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MailMerge {
    public static void main(String[] args) throws Exception {

        String output = "mailMerge.docx";

        // Create a Document instance
        Document document = new Document();

        // Add a section
        Section section = document.addSection();

        // Add 3 paragraphs to the section
        Paragraph para = section.addParagraph();
        Paragraph para2 = section.addParagraph();
        Paragraph para3 = section.addParagraph();

        // Add mail merge templates to each paragraph
        para.setText("Contact Name: ");
        para.appendField("Contact Name", FieldType.Field_Merge_Field);
        para2.setText("Phone: ");
        para2.appendField("Phone", FieldType.Field_Merge_Field);
        para3.setText("Date: ");
        para3.appendField("Date", FieldType.Field_Merge_Field);

        // Set the value for the mail merge template by the field name
        Date currentTime = new Date();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateString = formatter.format(currentTime);
        String[] filedNames = new String[] { "Contact Name", "Phone", "Date" };
        String[] filedValues = new String[] { "Robert Chan", "+1 (69) 123456", dateString };

        // Merge the specified value into template
        

        // save the document to file
        document.saveToFile(output, FileFormat.Docx);
    }
}