package pack;

import java.sql.*;

public class MysqlConnect{
	public static void main(String[] args) {
		System.out.println("MySQL Connect Example.");
		Connection conn = null;
		String url = "jdbc:mysql://localhost:3306/";
		String dbName = "jdbctutorial";
		String driver = "com.mysql.jdbc.Driver";
		String userName = "root"; 
		String password = "root";
		try {
			System.out.println("In try user 2 change");
			Class.forName(driver).newInstance();
			conn = DriverManager.getConnection(url+dbName,userName,password);
			System.out.println("Connected to the database");
			conn.close();
			System.out.println("Disconnected from database");
		} catch (Exception e) {
			System.out.println("Exception : "+e.getMessage());
			e.printStackTrace();
		}
	}
}