package Conexion;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;

public class DBConexion {

    private static String DRIVER = "org.mariadb.jdbc.Driver";
    private static String USR = "wmm";
    private static String PASS = "Hamrorc4";
    private static String URL = "jdbc:mariadb://joomla.aemet.es:3306/wmm_dtmad";
    static Connection con = null;

    static{
        try{
            Class.forName(DRIVER);
        }catch(ClassNotFoundException e){
            JOptionPane.showMessageDialog(null, "Drivers no encontrados" + e);

        }
    }

    public Connection getConnection(){

        try{
            con = DriverManager.getConnection(URL, USR, PASS);
            JOptionPane.showMessageDialog(null, "Conexión establecida");

        }catch(Exception e){
            JOptionPane.showMessageDialog(null, "Conexion no encontrada " + e);

        }

        return con;
    }

    public static ArrayList<String[]> getData(String tipo) throws SQLException {

        ArrayList<String[]> data = new ArrayList<>();

        if(con != null){

            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("SELECT * FROM almacen WHERE tipo = '" + tipo + "'");

//            if(rs != null){
//                System.out.println("consulta realizada");
//            }

            while(rs.next()){

                String [] fila = {rs.getString("ID"), rs.getString("etiqueta_AEMet"),
                        rs.getString("denominación"), rs.getString("cod_fabricante"), rs.getString("cantidad"), rs.getString("fecha_rec"),
                        rs.getString("fecha_mod"), rs.getString("tipo")};
                data.add(fila);

//                System.out.println("Fila 1: " + Arrays.toString(fila));
            }

            st.close();

        }else{

            JOptionPane.showMessageDialog(null, "Conexion no encontrada");
        }

        return data;
    }

    public void updateTabla(int fila,int columna, String nuevoValor, String tipo){

        PreparedStatement pst = null;

        try{

            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("SELECT column_name FROM information_schema.columns" +
                    " WHERE table_name = 'almacen' AND ordinal_position = " + (columna + 1) + ";");

            Statement st2 = con.createStatement();
            ResultSet rs2 = st2.executeQuery("SELECT * FROM almacen WHERE id = " + fila);

            rs.next();

            if(rs2.next()){

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            } else {

                Statement st3 = con.createStatement();
                st3.executeUpdate("INSERT INTO almacen VALUES (" + fila + ", null, null, null, null, null, null, '" + tipo + "')");

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            }

        }catch (SQLException e){

            JOptionPane.showMessageDialog(null, "Formato invalido");
            System.err.println(e);
//            e.printStackTrace();

        } finally {
            // Cerrar la conexión y el statement
            try {
                if (pst != null) pst.close();

            } catch (SQLException ex) {
                ex.printStackTrace();
            }
        }
    }

    public static void eliminarRegistro(int fila) throws SQLException {

        PreparedStatement pst = con.prepareStatement("DELETE FROM almacen WHERE id = " + fila);
        pst.executeUpdate();
    }

    public static void sobreescribirTabla(String tipo, DefaultTableModel modelo) throws SQLException {

        // Borramos los campos de este tipo
        PreparedStatement pst = con.prepareStatement("DELETE FROM almacen WHERE tipo = '" + tipo + "'");
        pst.executeUpdate();

        // Creamos los campos necesarios del mismo tipo
        ArrayList<String[]> row = new ArrayList<>();
        for(int f = 0 ; f < modelo.getRowCount() ; f++){

            ArrayList<String> campo = new ArrayList<>();

            for(int c = 0 ; c < modelo.getColumnCount()-1 ; c++){

                if(modelo.getValueAt(f, c) != null) campo.add(modelo.getValueAt(f, c).toString());
                else campo.add(null);

                System.out.println(modelo.getValueAt(f, c));
            }

            String k[] = campo.toArray(new String[campo.size()]);
            row.add(k);
        }

        // Añadimos los campos creados
        for(String[] fila : row){

            System.out.println("INSERT INTO almacen VALUES ('" + fila[0] + "', '" + fila[1] + "', '" + fila[2] + "', '" + fila[3] + "', '" + fila[4] + "', '" + fila[5] + "', '" + fila[6] + "', '" + tipo + "')");
            Statement st = con.createStatement();
            st.executeUpdate("INSERT INTO almacen VALUES ('" + fila[0] + "', " + fila[1] + ", '" + fila[2] + "', '" + fila[3] + "', " + fila[4] + ", " + fila[5] + ", " + fila[6] + ", '" + tipo + "')");
        }

        JOptionPane.showMessageDialog(null, "Base de Datos actualizada correctamente");
    }
}
