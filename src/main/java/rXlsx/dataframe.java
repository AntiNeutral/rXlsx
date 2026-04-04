package rXlsx;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZoneId;
import java.util.*;

public class dataframe {
    public ArrayList<String> columns;
    public ArrayList<String> column_type;
    public ArrayList<series> data;
    public int[] dim = {0, 0};

    public dataframe(ArrayList<String> column_name, ArrayList<String> column_type) {
        this.columns = column_name;
        this.column_type = column_type;
        this.dim[1] = column_name.size();
        this.data = new ArrayList<>();
    }

    public dataframe() {
        this.columns = new ArrayList<>();
        this.column_type = new ArrayList<>();
        this.data = new ArrayList<>();
    }

    private void extractColumnNames(Row row) {
        int size = row.getLastCellNum();
        this.dim[1] = size;
        for (int i = 0; i < size; i++) {
            Cell cell = row.getCell(i);
            if (!isNull(cell)) {
                this.columns.add(cell.getStringCellValue());
            } else {
                this.columns.add("column " + i);
            }
        }
    }

    public void defaultColumnType() {
        int diff = this.dim[1] - this.column_type.size();
        for (int i = 0; i < diff; i++) {
            this.column_type.add("str");
        }
    }

    public void setColumnType(int index, String type) {
        if (index <= this.dim[1] - 1) {
            this.column_type.set(index, type);
        } else {
            int diff = index - this.dim[1];
            for (int i = 0; i < diff; i++) {
                this.column_type.add("str");
            }
            this.column_type.add(type);
            this.dim[1] = index + 1;
        }
    }

    /**
     * Appends the given series to the dataframe and updates the dataframe dimensions. It does not handle the index.
     *
     * @param s the series to be appended to the dataframe
     */
    public void append(series s) {
        this.data.add(s);
        this.dim[0]++;
    }

    /**
     * Appends a list of series to the dataframe and updates the dataframe dimensions.
     * It does not handle the index.
     *
     * @param s the list of series to be appended to the dataframe
     */
    public void append(ArrayList<series> s) {
        this.data.addAll(s);
        this.dim[0] += s.size();
    }

    public void reindex() {
    }

    private static boolean isNull(Cell c) {
        if (c == null) {
            return true;
        } else if (c.getCellType() == CellType.BLANK) {
            return true;
        } else return c.getCellType() == CellType.STRING && c.getStringCellValue().trim().isEmpty();
    }

    private LinkedHashMap<String, Object> getRow(Row row) {
        LinkedHashMap<String, Object> row_data = new LinkedHashMap<>();
        for (int i = 0; i < this.dim[1]; i++){
            if (isNull(row.getCell(i))) {
                row_data.put(this.columns.get(i), null);
            } else {
                switch (this.column_type.get(i)) {
                    case "str":
                        row_data.put(this.columns.get(i), row.getCell(i).getStringCellValue());
                        break;
                    case "d":
                        row_data.put(this.columns.get(i), row.getCell(i).getNumericCellValue());
                        break;
                    case "int":
                        row_data.put(this.columns.get(i), (int) row.getCell(i).getNumericCellValue());
                        break;
                    case "date":
                        row_data.put(this.columns.get(i), row.getCell(i).getDateCellValue().toInstant().atZone(ZoneId.systemDefault())
                                .toLocalDate());
                        break;
                }
            }
        }
        return row_data;
    }

    public void readXLSX(String filePath, int page, boolean header) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(Files.newInputStream(Path.of(filePath)))) {
            Sheet sheet = workbook.getSheetAt(page);
            if (header) {
                Row header_row = sheet.getRow(0);
                if (this.columns.isEmpty()) {
                    this.extractColumnNames(header_row);
                    this.defaultColumnType();
                } else {
                    if (header_row.getLastCellNum() != this.dim[1]) {
                    throw new IllegalArgumentException("The number of columns in the sheet does not match the number of columns in the dataframe");
                    }
                }
                int row_num = sheet.getLastRowNum();
                for (int i = 1; i <= row_num; i++) {
                    LinkedHashMap<String, Object> row_data = this.getRow(sheet.getRow(i));
                    this.append(new series(this.dim[0], row_data));
                }
            } else if (!this.columns.isEmpty()){
                for (Row row: sheet){
                    LinkedHashMap<String, Object> row_data = this.getRow(row);
                    this.append(new series(this.dim[0], row_data));
                }
            } else {
                throw new IllegalArgumentException("The dataframe must have column names to read from an Excel file");
            }
        }
    }

    public void readXLSX(String filePath, int page) throws IOException {
        this.readXLSX(filePath, page, true);
    }
}