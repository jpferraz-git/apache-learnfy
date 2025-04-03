package school.sptech;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;

public class Main {

    public static void main(String[] args) throws IOException {
        String nomeArquivo = "teste.xlsx";

        Path caminho = Path.of(nomeArquivo);
        InputStream arquivo = Files.newInputStream(caminho);


        try {
            System.out.printf("\nIniciando leitura do arquivo %s\n%n", nomeArquivo);

            Workbook workbook = new XSSFWorkbook(arquivo);
            Sheet sheet = workbook.getSheetAt(0);

            int linhas = sheet.getLastRowNum();
            int colunas = sheet.getRow(1).getLastCellNum();

            for (int i = 0; i <= linhas; i++) {
                Row linhaAtual = sheet.getRow(i);

                for (int j = 0; j < colunas; j++) {
                    Cell celulaAtual = linhaAtual.getCell(j);
                    if (j == 0 && i == 1) {
                        System.out.println("--------------------------------------------------------------------------");
                    }
                    if (celulaAtual != null) {
                        switch (celulaAtual.getCellType()) {
                            case STRING: System.out.printf("| %s |", celulaAtual.getStringCellValue());
                                break;
                            case NUMERIC: System.out.printf("| %.0f |", celulaAtual.getNumericCellValue());
                                break;
                            case BOOLEAN: System.out.printf("| %-5b |", celulaAtual.getBooleanCellValue());
                                break;
                        }
                    }else{
                        System.out.print("| VALOR NULO |");
                    }
                }

                System.out.println();
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        arquivo.close();
    }

    private LocalDate converterDate(Date data) {
        return data.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
    }
