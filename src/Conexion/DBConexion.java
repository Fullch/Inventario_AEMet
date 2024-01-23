package Conexion;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;

public class DBConexion {

    private static String DRIVER = "com.mysql.cj.jdbc.Driver";
    private static String USR = "remote";
    private static String PASS = "h@R2P#tT!sYx7Z";
    private static String URL = "jdbc:mysql://172.24.160.53:3306/dietas/inventario";
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
            JOptionPane.showMessageDialog(null, "Conexion no encontrada" + e);

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

//            System.out.println(fila+1);
//            String id = String.valueOf(Integer.parseInt(rs2.getString(1)) - 1);

            if(rs2.next()){

//            System.out.println(id);

//                System.out.println(fila + " " + rs.getString("column_name") + " " + nuevoValor );

//            pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = " + id);
//            pst.setString(1, nuevoValor);
//
//            pst.executeUpdate();

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            } else {

                Statement st3 = con.createStatement();
                st3.executeUpdate("INSERT INTO almacen VALUES (" + fila + ", null, null, null, null, null, null, '" + tipo + "')");

//                System.out.println(fila);

//            System.out.println(id);

//                System.out.println(id + " " + rs.getString("column_name") + " " + nuevoValor );

//            pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = " + id);
//            pst.setString(1, nuevoValor);
//
//            pst.executeUpdate();

                pst = con.prepareStatement("UPDATE almacen SET " + rs.getString("column_name") + " = ? WHERE ID = "
                        + fila);
                pst.setString(1, nuevoValor);

                pst.executeUpdate();

            }

        }catch (SQLException e){

            e.printStackTrace();

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
        for(int f = 0 ; f < modelo.getRowCount()-1 ; f++){

            ArrayList<String> campo = new ArrayList<>();

            for(int c = 0 ; c < modelo.getColumnCount()-1 ; c++){

                if(modelo.getValueAt(f, c) != null) campo.add(modelo.getValueAt(f, c).toString());
                else campo.add(null);

//                System.out.println(modelo.getValueAt(f, c));
            }

            String k[] = campo.toArray(new String[campo.size()]);
            row.add(k);
        }

        // Añadimos los campos creados
        for(String[] fila : row){

//            System.out.println("INSERT INTO almacen VALUES ('" + fila[0] + "', '" + fila[1] + "', '" + fila[2] + "', '" + fila[3] + "', '" + fila[4] + "', '" + fila[5] + "', '" + fila[6] + "', '" + tipo + "')");
            Statement st = con.createStatement();
            st.executeUpdate("INSERT INTO almacen VALUES ('" + fila[0] + "', " + fila[1] + ", '" + fila[2] + "', '" + fila[3] + "', " + fila[4] + ", " + fila[5] + ", " + fila[6] + ", '" + tipo + "')");
        }

        JOptionPane.showMessageDialog(null, "Base da Datos actualizada correctamente");
    }
}
