
import java.util.List;


public class ExcelDataRowModel {
    private int rowIndex;
    private List<Object> dataRows;

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public List<Object> getDataRows() {
        return dataRows;
    }

    public void setDataRows(List<Object> dataRows) {
        this.dataRows = dataRows;
    }
}
