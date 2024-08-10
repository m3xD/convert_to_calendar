import controller.Calendar;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        String path = "src/ThoiKhoaBieuSinhVien.xls";
        Calendar s = new Calendar();
        s.processExcel(path, "F11", "F30");
    }
}