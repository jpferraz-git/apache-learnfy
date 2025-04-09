import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;


import java.io.InputStream;
import org.apache.poi.util.IOUtils;

    public static void main(String[] args) throws Exception{
        IOUtils.setByteArrayMaxOverride(1_000_000_000);

        OPCPackage arquivo = OPCPackage.open("vai-explodir.xlsx");

        XSSFReader leitor = new XSSFReader(arquivo);

        ReadOnlySharedStringsTable texto = new ReadOnlySharedStringsTable(arquivo);

        InputStream folha = leitor.getSheetsData().next();

        SheetContentsHandler manipulador = new SheetContentsHandler() {

            @Override
            public void startRow(int rowNum) {
                System.out.print("Linha " + rowNum + ": ");
            }

            @Override
            public void endRow(int rowNum) {
                System.out.println();
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {

                double stringInteira = 0;
                if (formattedValue.matches("^[0-9].*")) {
                    stringInteira = Double.parseDouble(formattedValue);
                    System.out.print(stringInteira + " | ");

                } else if (formattedValue == null) {
                    System.out.print(" Valor Nulo | ");

                } else {
                    System.out.print(formattedValue + " | ");
                }
            }
        };

        XMLReader parser = XMLReaderFactory.createXMLReader();
        XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(null, null, texto, manipulador, null, false);
        parser.setContentHandler(xmlHandler);
        parser.parse(new InputSource(folha));
        folha.close();

}
