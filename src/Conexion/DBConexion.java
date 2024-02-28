package Conexion;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.sql.*;
import java.util.ArrayList;

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

            while(rs.next()){

                String [] fila = {rs.getString("ID"), rs.getString("etiqueta_AEMet"),
                        rs.getString("denominación"), rs.getString("cod_fabricante"), rs.getString("cantidad"), rs.getString("fecha_rec"),
                        rs.getString("fecha_mod"), rs.getString("detalles"), rs.getString("tipo")};
                data.add(fila);

            }

            st.close();

        }else{

            JOptionPane.showMessageDialog(null, "Conexion no encontrada");
        }

        return data;
    }

    public static ArrayList<String> iniciarTipos() throws SQLException {

        ArrayList<String> tipos = new ArrayList<>();

        Statement st = con.createStatement();
        ResultSet rs = st.executeQuery("SELECT DISTINCT tipo FROM almacen;");

        while(rs.next()){

            tipos.add(rs.getString(1));
        }

        return tipos;
    }

    public static void añadirFilas(int it, String tipo, DefaultTableModel modelo) throws SQLException {

        ArrayList<String[]> row = new ArrayList<>();
        it++;
        for(int f = 0 ; f < modelo.getRowCount() ; f++){

            ArrayList<String> campos = new ArrayList<>();

            campos.add(String.valueOf(it++));

            for(int c = 1 ; c < modelo.getColumnCount()-1 ; c++){

                if(modelo.getValueAt(f, c) != null) campos.add(modelo.getValueAt(f, c).toString());
                else campos.add(null);

//                System.out.println("Añadir campo " + modelo.getValueAt(f, c));
            }

            String[] k = campos.toArray(new String[campos.size()]);
            row.add(k);
        }

        for(String[] fila : row){

            PreparedStatement ps = con.prepareStatement("INSERT INTO almacen VALUES (?, ?, ?, ?, ?, ?, ?, ?, '" + tipo + "')");
            for(int i = 0 ; i < fila.length ; i++){
                ps.setString(i+1, fila[i]);
            }
            ps.executeUpdate();

        }
    }

    public void updateTabla(int fila,int columna, String nuevoValor, String tipo){

        PreparedStatement pst = null;

        try{

            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery("SELECT column_name FROM information_schema.columns" +
                    " WHERE table_name = 'almacen' AND ordinal_position = " + (columna + 1) + ";");

            Statement st2 = con.createStatement();
            ResultSet rs2 = st2.executeQuery("SELECT * FROM almacen WHERE id = " + fila + " AND tipo = '" + tipo + "'");

            rs.next();

            if(rs2.next()){

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            } else {

                Statement st3 = con.createStatement();
                st3.executeUpdate("INSERT INTO almacen VALUES (" + fila + ", null, null, null, null, null, null, null, '" + tipo + "')");

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            }

        }catch (SQLException e){

            JOptionPane.showMessageDialog(null, "Formato invalido");
            e.printStackTrace();

        } finally {

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

    public static void sobreescribirTabla(int it, String tipo, DefaultTableModel modelo) throws SQLException {

        // Borramos los campos de este tipo
        PreparedStatement pst = con.prepareStatement("DELETE FROM almacen WHERE tipo = '" + tipo + "'");
        pst.executeUpdate();

        // Creamos los campos necesarios del mismo tipo
        ArrayList<String[]> row = new ArrayList<>();
        it++;

        for(int f = 0 ; f < modelo.getRowCount() ; f++){

            ArrayList<String> campos = new ArrayList<>();

            campos.add(String.valueOf(it++));
            System.out.println("Añadir filas " + it);

            for(int c = 1 ; c < modelo.getColumnCount()-1 ; c++){

                if(modelo.getValueAt(f, c) != null) campos.add(modelo.getValueAt(f, c).toString());
                else campos.add(null);

            }

            String k[] = campos.toArray(new String[campos.size()]);
            row.add(k);
        }

        // Añadimos los campos creados
        for(String[] fila : row){

//            System.out.println("INSERT INTO almacen VALUES ('" + fila[0] + "', '" + fila[1] + "', '" + fila[2] + "', '" + fila[3] + "', '" + fila[4] + "', '" + fila[5] + "', '" + fila[6] + "', '" + tipo + "')");
            PreparedStatement ps = con.prepareStatement("INSERT INTO almacen VALUES (?, ?, ?, ?, ?, ?, ?, ?, '" + tipo + "')");
            for(int i = 0 ; i < fila.length ; i++){

                ps.setString(i+1, fila[i]);
            }
            ps.executeUpdate();

        }

        JOptionPane.showMessageDialog(null, "Base de Datos actualizada correctamente");
    }
}
