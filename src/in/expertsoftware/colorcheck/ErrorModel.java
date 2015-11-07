package in.expertsoftware.colorcheck;

public class ErrorModel {
	
	String sheet_name;
        String error_level;
	String error_desc;
	String cell_ref;
	int row;
	int col;
        public String getError_level() {
		return error_level;
	}
	public void setError_level(String error_level) {
		this.error_level = error_level;
        }
        public String getCell_ref() {
		return cell_ref;
	}
	public void setCell_ref(String cell_ref) {
		this.cell_ref = cell_ref;
	}
	public String getSheet_name() {
		return sheet_name;
	}
	public void setSheet_name(String sheet_name) {
		this.sheet_name = sheet_name;
	}
	public String getError_desc() {
		return error_desc;
	}
	public void setError_desc(String error_desc) {
		this.error_desc = error_desc;
	}
	public int getRow() {
		return row;
	}
	public void setRow(int row) {
		this.row = row;
	}
	public int getCol() {
		return col;
	}
	public void setCol(int col) {
		this.col = col;
	}
}