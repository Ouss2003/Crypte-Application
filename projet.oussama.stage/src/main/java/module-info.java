module com.example.demo {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.pdfbox;
    requires org.apache.poi.ooxml;


    opens com.example.demo1 to javafx.fxml;
    exports com.example.demo1;
}