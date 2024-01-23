import Conexion.DBConexion;
import UI.GUI;
import UI.Interfaz;

import java.sql.SQLException;

public class Main {
    public static void main(String[] args) throws SQLException {

        DBConexion con = new DBConexion();

        if(con.getConnection() != null){

            new GUI(con);
        }
    }
}