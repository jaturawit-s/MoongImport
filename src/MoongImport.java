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
		
		// Test
		ds.setServerName("122.155.10.96");
		ds.setPortNumber(Integer.parseInt("1390"));
		
		// Production
//		ds.setServerName("localhost");
//		ds.setPortNumber(Integer.parseInt("1390"));
		
		ds.setDatabaseName("MoongDB");
		
        try (Connection connection =  ds.getConnection();
        		Statement stmt = connection.createStatement();) {
        	
        	
        	String SQL = "";
        	
        	// NetsalesGTbyCustomerbySKU.xlsx
        	SQL = "IF OBJECT_ID('dbo.NetsalesGTbyCustomerbySKU', 'U') IS NOT NULL Drop Table dbo.NetsalesGTbyCustomerbySKU";
        	stmt.execute(SQL);
        	
        	SQL = "SELECT * INTO dbo.NetsalesGTbyCustomerbySKU " + 
        			"FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0; Database=E:\\FTP\\NetsalesGTbyCustomerbySKU.xlsx', [Sheet1$]);";
            stmt.execute(SQL);
            
            System.out.println("NetsalesGTbyCustomerbySKU.xlsx - done");
            
            // DummyHistoryNetsales.xlsx
        	SQL = "IF OBJECT_ID('dbo.DummyHistoryNetsales', 'U') IS NOT NULL Drop Table dbo.DummyHistoryNetsales";
        	stmt.execute(SQL);
        	
        	SQL = "SELECT * INTO dbo.DummyHistoryNetsales " + 
        			"FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0; Database=E:\\FTP\\DummyHistoryNetsales.xlsx', [Sheet1$]);"; 
            stmt.execute(SQL);
            
            System.out.println("DummyHistoryNetsales.xlsx - done");
            
            // CustomerMaster.xlsx        
            SQL = "IF OBJECT_ID('dbo.CustomerMaster', 'U') IS NOT NULL Drop Table dbo.CustomerMaster";
        	stmt.execute(SQL);
        	
        	SQL = "SELECT * INTO dbo.CustomerMaster " + 
        			"FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0; Database=E:\\FTP\\CustomerMaster.xlsx', [Sheet1$]);"; 
            stmt.execute(SQL);
            
            System.out.println("CustomerMaster.xlsx - done");
            
            // Target.xlsx
            SQL = "IF OBJECT_ID('dbo.Target', 'U') IS NOT NULL Drop Table dbo.Target";
        	stmt.execute(SQL);
        	
        	SQL = "SELECT * INTO dbo.Target " + 
        			"FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0; Database=E:\\FTP\\Target.xlsx', [Sheet1$]);"; 
            stmt.execute(SQL);
            
            System.out.println("Target.xlsx - done");
        }
        // Handle any errors that may have occurred.
        catch (SQLException e) {
            e.printStackTrace();
        }
    }
}