package school.sptech;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeitorExcel {

    public void extrarLivros(String nomeArquivo, InputStream arquivo) {
        try {
            System.out.println("\nIniciando leitura do arquivo %s\n".formatted(nomeArquivo));

            // Criando um objeto Workbook a partir do arquivo recebido
            Workbook workbook = new XSSFWorkbook(arquivo);

            Sheet sheet = workbook.getSheetAt(0);

                int linhas = sheet.getLastRowNum();
                int colunas = sheet.getRow(1).getLastCellNum();

            for (int i = 0; i <= linhas; i++) {
                Row linhaAtual = sheet.getRow(i);

                for (int j = 0; j < colunas; j++) {
                    Cell celulaAtual = linhaAtual.getCell(j);
                    if (j == 0 && i == 1){
                        System.out.println("--------------------------------------------------------------------------");
                    }
                    switch (celulaAtual.getCellType()){
                        case STRING: System.out.printf("| %s |", celulaAtual.getStringCellValue()); break;
                        case NUMERIC: System.out.printf("| %.0f |", celulaAtual.getNumericCellValue()); break;
                        case BOOLEAN: System.out.printf("| %-5b |",celulaAtual.getBooleanCellValue()); break;
                    }
                }

                System.out.println();
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private LocalDate converterDate(Date data) {
        return data.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
}
