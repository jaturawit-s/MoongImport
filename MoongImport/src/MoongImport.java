import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import com.microsoft.sqlserver.jdbc.SQLServerDataSource;

public class MoongImport {

	// Connect to your database.
	public static void main(String[] args) {

		SQLServerDataSource ds = new SQLServerDataSource();
		ds.setUser("sa");
		ds.setPassword("Password$1");

		String[] file = new String[6];
		String env = "production";

		file[0] = "CustomerMaster";
		file[1] = "DummyHistoryNetsales";
		file[2] = "NetsalesbyCustomerbySKU";
		file[3] = "Target";
		file[4] = "TargetRatioByRegion";
		file[5] = "TargetRatioBySalesman";

		if (env.equals("production")) {
			// Test
			ds.setServerName("122.155.10.96");
			ds.setPortNumber(Integer.parseInt("1390"));
		} else if (env.equals("test")) {
			// Production
			ds.setServerName("localhost");
			ds.setPortNumber(Integer.parseInt("1390"));
		}

		ds.setDatabaseName("MoongDB");

		try (Connection connection = ds.getConnection(); Statement stmt = connection.createStatement();) {

			for (int i = 0; i < file.length; i++) {

				String SQL = "";

				SQL = "IF OBJECT_ID('dbo." + file[i] + "', 'U') IS NOT NULL Drop Table dbo." + file[i];
				stmt.execute(SQL);

				SQL = "SELECT * INTO dbo." + file[i]
						+ " FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0; Database=D:\\FTP\\" + file[i]
						+ ".xlsx', [Sheet1$]);";
				stmt.execute(SQL);

				System.out.println(file[i] + ".xlsx - done");

				stmt.executeQuery(
						"Select Count(*) From INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME='" + "dbo." + file[i] + "'");
				

			}
		}
		// Handle any errors that may have occurred.
		catch (SQLException e) {
			e.printStackTrace();
		}
	}
}