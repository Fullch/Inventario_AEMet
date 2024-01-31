package UI;

import Conexion.DBConexion;
import com.intellij.uiDesigner.core.GridConstraints;
import com.intellij.uiDesigner.core.GridLayoutManager;
import net.miginfocom.swing.MigLayout;
import net.sourceforge.jdatepicker.JDatePicker;
import net.sourceforge.jdatepicker.impl.JDatePanelImpl;
import net.sourceforge.jdatepicker.impl.JDatePickerImpl;
import net.sourceforge.jdatepicker.impl.UtilDateModel;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.event.TableModelEvent;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.plaf.FontUIResource;
import javax.swing.table.*;
import javax.swing.text.*;
import java.awt.*;
import java.awt.Font;
import java.io.*;
import java.sql.SQLException;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.logging.Logger;

// TODO:
//  -Crear plantilla para importar tabla
//  -Habilitar y arreglar botones

public class Otra extends JFrame {

    private JTabbedPane pestanas;
    private JTextField campoBusqueda;
    private JButton botonExportar;
    private JButton botonImportar;
    private JButton botonExportarT;
    private JButton botonPrueba;
    private JPanel panelPrincipal;
    private JPanel panelBotones;
    private JPanel panelBusqueda;

    static int it = 0;
    static String[] tipos = {"SSBB", "Informática", ""};
    static ArrayList<DefaultTableModel> modeloTablas = new ArrayList<>();
    static ArrayList<JTable> tablas = new ArrayList<>();

    public Otra(DBConexion con) throws SQLException {

        $$$setupUI$$$();

        // Configuración de la ventana
        setTitle("INVENTARIO");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(1000, 1000);
        setLocationRelativeTo(null);
        setVisible(true);
        setResizable(false);

        crearTablas(tipos, modeloTablas, con);

        campoBusqueda.addActionListener(e -> {

            RowFilter<TableModel, Integer> rowFilter = RowFilter.regexFilter(campoBusqueda.getText());

            int selectedIndex = pestanas.getSelectedIndex();

            TableRowSorter<TableModel> sorter = new TableRowSorter<>();
            sorter = (TableRowSorter<TableModel>) tablas.get(selectedIndex).getRowSorter();
            sorter.setRowFilter(rowFilter);

        });

        // Botón que exporta a Excel
        botonExportar.setFont(new Font("Arial", Font.PLAIN, 16));
        botonExportar.setPreferredSize(new Dimension(150, 50));

        botonExportar.addActionListener(e -> {

            int selectedIndex = pestanas.getSelectedIndex();

            if(selectedIndex != -1){

                try {
                    exportarExcel(tablas.get(selectedIndex));
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }

        });

        // Botón que importa a Excel
        botonImportar.setFont(new Font("Arial", Font.PLAIN, 16));
        botonImportar.setPreferredSize(new Dimension(150, 50));
//        panelBotones.add(botonImportar);

        botonImportar.addActionListener(e -> {

            int selectedIndex = pestanas.getSelectedIndex();

            if(selectedIndex != -1){

                try {
                    importarExcel(tablas.get(selectedIndex));
                } catch (SQLException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });

        // Botón que exporta toda la tabla a Excel
        botonExportarT.setFont(new Font("Arial", Font.PLAIN, 16));
        botonExportarT.setPreferredSize(new Dimension(150, 50));
//        panelBotones.add(botonExportarT);

        botonExportarT.addActionListener(e -> {

            try {
                exportarTodo(tipos);
            } catch (IOException | SQLException ex) {
                throw new RuntimeException(ex);
            }

        });

        // Botón de prueba
        botonPrueba.setFont(new Font("Arial", Font.PLAIN, 16));
        botonPrueba.setPreferredSize(new Dimension(150, 50));


        botonPrueba.addActionListener(e -> {

        });

        // Agregar el panel principal a la ventana
        add(panelPrincipal);

        // Una última vez para que queden bien antes de poder interactuar, después se realizará
        // de forma automática
        it = refrescarTablas(modeloTablas, tipos, it);

        // Hacer visible la ventana
        setVisible(true);
    }

    private void crearTablas(String[] tipos, ArrayList<DefaultTableModel> modeloTablas, DBConexion con) {

        for(int i = 0 ; i < tipos.length ; i++){

            TableRowSorter<TableModel> sorter;
            DefaultTableModel modeloTabla = new DefaultTableModel(){

                @Override
                public boolean isCellEditable(int row, int column) {
                    if(column == 0) return false;
                    return true;
                }

            };
            modeloTablas.add(modeloTabla);

            try {

                modeloTabla.addColumn("ID");
                modeloTabla.addColumn("Etiqueta AEMet");
                modeloTabla.addColumn("Denominación");
                modeloTabla.addColumn("Código Fabricante");
                modeloTabla.addColumn("Cantidad");
                modeloTabla.addColumn("Fecha Recepción");
                modeloTabla.addColumn("Fecha Modificación");
                modeloTabla.addColumn("");

                JTable tabla = new JTable(modeloTabla);
                tabla.setFont(new Font("", Font.PLAIN, 16));
                tabla.getTableHeader().setFont(new Font("", Font.BOLD, 16));

                tabla.getColumnModel().getColumn(1).setCellEditor(new NumberEditor());
                tabla.getColumnModel().getColumn(4).setCellEditor(new NumberEditor());

                tabla.getColumnModel().getColumn(3).setCellEditor(new CustomStringEditor());

                CalendarioRenderer cr = new CalendarioRenderer();
                CalendarioEditor ce = new CalendarioEditor();

                tabla.getColumnModel().getColumn(5).setCellRenderer(cr);
                tabla.getColumnModel().getColumn(5).setCellEditor(ce);
                tabla.getColumnModel().getColumn(6).setCellRenderer(cr);
                tabla.getColumnModel().getColumn(6).setCellEditor(ce);

                tabla.getColumnModel().getColumn(7).setCellRenderer(new ButtonRenderer());
                tabla.getColumnModel().getColumn(7).setCellEditor(new ButtonEditor(tabla));

                sorter = new TableRowSorter<>(modeloTabla);
                tabla.setRowSorter(sorter);

                tabla.setRowHeight(25);

                tablas.add(tabla);

                it = refrescarTablas(modeloTablas, tipos, it);

                JScrollPane scrollPane = new JScrollPane(tabla);
                JPanel panelTabla = new JPanel(new MigLayout("fill"));
                panelTabla.setLayout(new MigLayout("wrap, fill"));
                panelTabla.add(scrollPane, "grow");

                pestanas.addTab(tipos[i], panelTabla);

                tabla.repaint();

                modeloTabla.addTableModelListener(e -> {

                    if (e.getType() == TableModelEvent.UPDATE) {

                        try{
                            int fila = e.getFirstRow();
                            int columna = e.getColumn();
                            int id = Integer.parseInt((String) tabla.getValueAt(fila, 0));

                            if (columna != 6) {

                                if(columna == 3){


                                }

                                String nuevoValor;

                                if(modeloTabla.getValueAt(fila, columna) != null){

                                    nuevoValor = modeloTabla.getValueAt(fila, columna).toString();

                                }else {

                                    nuevoValor = null;
                                }

                                con.updateTabla(id, columna, nuevoValor, tipos[0]);

                                try {
                                    it = refrescarTablas(modeloTablas, tipos, it);
                                } catch (SQLException ex) {
                                    throw new RuntimeException(ex);
                                }
                            }
                        }catch(Exception ex){

                            JOptionPane.showMessageDialog(null, "Formato invalido");
                            System.err.println(ex);
//                            ex.printStackTrace();
                        }
                    }

                });

            } catch (SQLException ex) {
                throw new RuntimeException(ex);
            }

        }
    }

    private static int refrescarTablas(ArrayList<DefaultTableModel> modeloTablas, String[] tipos, int it) throws SQLException {

        int i = 0;

        for (DefaultTableModel modeloTabla : modeloTablas) {

            it = refrescarTabla(modeloTabla, tipos[i], it);
            i++;
        }

        return it;
    }

    private static int refrescarTabla(DefaultTableModel modeloTabla, String tipo, int it) throws SQLException {

        modeloTabla.setRowCount(0);

        ArrayList<String[]> arrayListTabla = new ArrayList<>(DBConexion.getData(tipo));

        if (!arrayListTabla.isEmpty()) {

            String pattern = "yyyy-MM-dd";
            DateTimeFormatter inputFormat = DateTimeFormatter.ofPattern(pattern);

            for (String[] fila : arrayListTabla) {

                if (fila[5] != null) {

                    LocalDate parsed = LocalDate.parse(fila[5], inputFormat);
                    DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                    String formattedDate = parsed.format(outputFormatter);
                    fila[5] = formattedDate;
                }
                if (fila[6] != null) {

                    LocalDate parsed = LocalDate.parse(fila[6], inputFormat);
                    DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                    String formattedDate = parsed.format(outputFormatter);
                    fila[6] = formattedDate;
                }

                if (Integer.parseInt(fila[0]) > it) {
                    it = Integer.parseInt(fila[0]);
                }

                fila[7] = "";

                modeloTabla.addRow(fila);
            }

        } else {

            it = 0;
        }

        // Esto hace falta porque el addRow requiere un array aunque sea vacío
        ArrayList<String[]> arrayVacio = new ArrayList<>();
        String[] vacio = {it + 1 + ""};
        arrayVacio.add(vacio);
        modeloTabla.addRow(arrayVacio.getFirst());

        return it;
    }


    /**
     * Method generated by IntelliJ IDEA GUI Designer
     * >>> IMPORTANT!! <<<
     * DO NOT edit this method OR call it in your code!
     *
     * @noinspection ALL
     */
    private void $$$setupUI$$$() {
        panelPrincipal = new JPanel();
        panelPrincipal.setLayout(new GridLayoutManager(3, 1, new Insets(0, 0, 0, 0), -1, -1));
        pestanas = new JTabbedPane();
        Font pestanasFont = this.$$$getFont$$$(null, -1, 20, pestanas.getFont());
        if (pestanasFont != null) pestanas.setFont(pestanasFont);
        panelPrincipal.add(pestanas, new GridConstraints(2, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, null, new Dimension(200, 200), null, 0, false));
        panelBotones = new JPanel();
        panelBotones.setLayout(new GridLayoutManager(2, 3, new Insets(20, 0, 20, 0), -1, -1));
        panelPrincipal.add(panelBotones, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        botonExportar = new JButton();
        botonExportar.setText("Exportar");
        panelBotones.add(botonExportar, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        botonImportar = new JButton();
        botonImportar.setText("Importar");
        panelBotones.add(botonImportar, new GridConstraints(0, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        botonExportarT = new JButton();
        botonExportarT.setText("Exportar todo");
        panelBotones.add(botonExportarT, new GridConstraints(0, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        botonPrueba = new JButton();
        botonPrueba.setText("Probar");
        panelBotones.add(botonPrueba, new GridConstraints(1, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        panelBusqueda = new JPanel();
        panelBusqueda.setLayout(new GridLayoutManager(1, 1, new Insets(0, 0, 0, 0), -1, -1));
        panelPrincipal.add(panelBusqueda, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        campoBusqueda = new JTextField();
        Font campoBusquedaFont = this.$$$getFont$$$(null, -1, 20, campoBusqueda.getFont());
        if (campoBusquedaFont != null) campoBusqueda.setFont(campoBusquedaFont);
        panelBusqueda.add(campoBusqueda, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(251, 30), null, 0, false));
    }

    /**
     * @noinspection ALL
     */
    private Font $$$getFont$$$(String fontName, int style, int size, Font currentFont) {
        if (currentFont == null) return null;
        String resultName;
        if (fontName == null) {
            resultName = currentFont.getName();
        } else {
            Font testFont = new Font(fontName, Font.PLAIN, 10);
            if (testFont.canDisplay('a') && testFont.canDisplay('1')) {
                resultName = fontName;
            } else {
                resultName = currentFont.getName();
            }
        }
        Font font = new Font(resultName, style >= 0 ? style : currentFont.getStyle(), size >= 0 ? size : currentFont.getSize());
        boolean isMac = System.getProperty("os.name", "").toLowerCase(Locale.ENGLISH).startsWith("mac");
        Font fontWithFallback = isMac ? new Font(font.getFamily(), font.getStyle(), font.getSize()) : new StyleContext().getFont(font.getFamily(), font.getStyle(), font.getSize());
        return fontWithFallback instanceof FontUIResource ? fontWithFallback : new FontUIResource(fontWithFallback);
    }

    /**
     * @noinspection ALL
     */
    public JComponent $$$getRootComponent$$$() {
        return panelPrincipal;
    }

    static class NumberEditor extends DefaultCellEditor {

        public NumberEditor() {
            super(new JFormattedTextField());
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {

            JFormattedTextField editor = (JFormattedTextField) super.getTableCellEditorComponent(table, value, isSelected, row, column);

            if (value instanceof Number){
                Locale myLocale = Locale.getDefault();

                NumberFormat numberFormatB = NumberFormat.getInstance(myLocale);
                numberFormatB.setMaximumFractionDigits(2);
                numberFormatB.setMinimumFractionDigits(2);
                numberFormatB.setMinimumIntegerDigits(1);

                editor.setFormatterFactory(new javax.swing.text.DefaultFormatterFactory(
                        new NumberFormatter(numberFormatB)));

                editor.setHorizontalAlignment(SwingConstants.RIGHT);
                editor.setValue(value);
            }

            return editor;
        }

        @Override
        public boolean stopCellEditing() {

            return super.stopCellEditing();

        }

        @Override
        public Object getCellEditorValue() {
            // get content of textField
            String str = (String) super.getCellEditorValue();
            if (str == null) {
                return null;
            }

            if (str.length() == 0) {
                return null;
            }

            // try to parse a number
            try {
                ParsePosition pos = new ParsePosition(0);
                Number n = NumberFormat.getInstance().parse(str, pos);
                if (pos.getIndex() != str.length()) {
                    throw new ParseException(
                            "parsing incomplete", pos.getIndex());
                }

                // return an instance of column class
                return n.intValue();

            } catch (ParseException pex) {
                throw new RuntimeException(pex);
            }
        }

    }

    static class CalendarioRenderer extends DefaultTableCellRenderer {
        JDatePicker datePicker;

        public CalendarioRenderer() {
            UtilDateModel model = new UtilDateModel();
            JDatePanelImpl datePanel = new JDatePanelImpl(model);
            datePicker = new JDatePickerImpl(datePanel, new Interfaz.DateLabelFormatter());
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {

            if (value instanceof Date) {

                UtilDateModel model = new UtilDateModel();
                model.setValue((Date) value);

                // Crea un nuevo JDatePicker con el modelo
                return new JDatePickerImpl(new JDatePanelImpl(model), new Interfaz.DateLabelFormatter());

            } else {

                return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
            }
        }

    }

    static class CalendarioEditor extends AbstractCellEditor implements TableCellEditor {
        private final JDatePickerImpl datePicker;
        public Date selectedDate;

        public CalendarioEditor() {
            UtilDateModel model = new UtilDateModel();
            JDatePanelImpl datePanel = new JDatePanelImpl(model);
            datePicker = new JDatePickerImpl(datePanel, new Interfaz.DateLabelFormatter());

            datePicker.addActionListener(e -> stopCellEditing());
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {

            UtilDateModel model;
            if (value instanceof Date) {
                model = new UtilDateModel((Date) value);
            } else {
                model = new UtilDateModel();
            }
            datePicker.getModel().setDate(model.getYear(), model.getMonth(), model.getDay());

            return datePicker;
        }

        @Override
        public Object getCellEditorValue() {
            selectedDate = (Date) datePicker.getModel().getValue();
            String pattern = "yyyy-MM-dd";
            //            System.out.println(dateInString);
            return new SimpleDateFormat(pattern).format(selectedDate);
        }

    }

    // Clase de renderizador para mostrar el botón en la celda
    static class ButtonRenderer extends JButton implements TableCellRenderer {
        // Constructor
        public ButtonRenderer() {
            setOpaque(true);
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            setText((value == null) ? "" : value.toString());
            return this;
        }
    }

    // Clase de editor para manejar los eventos del botón en la celda
    static class ButtonEditor extends DefaultCellEditor {
        private final JButton button;
        private boolean isPushed;
        private final JTable table;

        public ButtonEditor(JTable table) {
            super(new JCheckBox());
            this.table = table;
            button = new JButton();
            button.setOpaque(true);
            button.addActionListener(e -> {

                try{
                    fireEditingStopped();
                }catch(Exception ex){
                    System.err.println("Hay una excepción que no es necesario arreglarla hasta donde yo sé: \n" + ex);
                }
            });
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {

            if (isSelected) {
                button.setForeground(table.getSelectionForeground());
                button.setBackground(table.getSelectionBackground());
            } else {
                button.setForeground(table.getForeground());
                button.setBackground(table.getBackground());
            }

            //Nombre del botón
            int id = Integer.parseInt((String) table.getValueAt(row, 0));
            button.setActionCommand("Botón_" + id + "_" + row + "_" + column);

            isPushed = true;
            return button;
        }

        @Override
        public Object getCellEditorValue() {

            if (isPushed) {
                // Manejar la acción del botón
                String actionCommand = button.getActionCommand();
                String[] parts = actionCommand.split("_");
                int id = Integer.parseInt(parts[1]);
                try {
                    DBConexion.eliminarRegistro(id);
                } catch (SQLException e) {
                    throw new RuntimeException(e);
                }

                int fila = Integer.parseInt(parts[2]);
                // Eliminar la fila correspondiente a la celda
                ((DefaultTableModel) table.getModel()).removeRow(fila);
                try {
                    it = refrescarTablas(modeloTablas, tipos, it);
                } catch (SQLException e) {
                    throw new RuntimeException(e);
                }
            }

            isPushed = false;
            return false;
        }

        @Override
        public boolean stopCellEditing() {
            isPushed = false;
            return super.stopCellEditing();
        }
    }

    public void exportarExcel(JTable tabla) throws IOException {

        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls");
        chooser.setFileFilter(filter);
        chooser.setDialogTitle("Guardar archivo");
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {

            String ruta = chooser.getSelectedFile().toString().concat(".xls");

            File archivoXLS = new File(ruta);
            if (archivoXLS.exists()) {
                archivoXLS.delete();
            }

            archivoXLS.createNewFile();

            HSSFWorkbook libro = new HSSFWorkbook();
            FileOutputStream archivo = new FileOutputStream(archivoXLS);
            Sheet hoja = libro.createSheet("Tabla");
            hoja.setDisplayGridlines(true);

            HSSFCellStyle estilo = libro.createCellStyle();
            HSSFFont font = libro.createFont();
            font.setBold(true);
            estilo.setFont(font);
            estilo.setWrapText(true);
            estilo.setBorderBottom(BorderStyle.valueOf((short) 1));
            estilo.setBorderLeft(BorderStyle.valueOf((short) 1));
            estilo.setBorderRight(BorderStyle.valueOf((short) 1));
            estilo.setBorderTop(BorderStyle.valueOf((short) 1));

            CellStyle estiloCelda = libro.createCellStyle();
            estiloCelda.setWrapText(true);
            estiloCelda.setBorderBottom(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderLeft(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderRight(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderTop(BorderStyle.valueOf((short) 1));

            for (int f = 0; f < tabla.getRowCount(); f++) {
                Row fila = hoja.createRow(f);
                for (int c = 0; c < tabla.getColumnCount(); c++) {
                    Cell celda = fila.createCell(c);
                    fila.setHeight((short) 450);
                    if (f == 0) {
                        celda.setCellStyle(estilo);
                        celda.setCellValue(tabla.getColumnName(c));
                    }
                }
            }

            int filaInicio = 1;
            for (int f = 0; f < tabla.getRowCount(); f++) {
                Row fila = hoja.createRow(filaInicio);
                fila.setHeight((short) 450);
                filaInicio++;
                for (int c = 0; c < tabla.getColumnCount() - 1; c++) {

                    Cell celda = fila.createCell(c);
                    celda.setCellStyle(estiloCelda);
                    hoja.setColumnWidth(c, 4500);

                    if (tabla.getValueAt(f, c) instanceof Double) {
                        celda.setCellValue(Double.parseDouble(tabla.getValueAt(f, c).toString()));
                    } else if (tabla.getValueAt(f, c) instanceof Float) {
                        celda.setCellValue(Float.parseFloat((String) tabla.getValueAt(f, c)));
                    } else if (tabla.getValueAt(f, c) instanceof String){
                        celda.setCellValue(String.valueOf(tabla.getValueAt(f, c)));
                    } else celda.setCellValue("");
                }
            }

            try {
                libro.write(archivo);
                archivo.close();
                Desktop.getDesktop().open(archivoXLS);
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null, "Error al abrir el archivo: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }

        }
    }

    @SuppressWarnings("ResultOfMethodCallIgnored")
    private void exportarTodo(String[] tipos) throws IOException, SQLException {

        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls");
        chooser.setFileFilter(filter);
        chooser.setDialogTitle("Guardar archivo");
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {

            String ruta = chooser.getSelectedFile().toString().concat(".xls");

            File archivoXLS = new File(ruta);
            if (archivoXLS.exists()) {
                archivoXLS.delete();
            }

            archivoXLS.createNewFile();
            HSSFWorkbook libro = new HSSFWorkbook();
            FileOutputStream archivo = new FileOutputStream(archivoXLS);
            Sheet hoja = libro.createSheet("TablaCompleta");
            hoja.setDisplayGridlines(true);

            HSSFCellStyle estilo = libro.createCellStyle();
            HSSFFont font = libro.createFont();
            font.setBold(true);
            estilo.setFont(font);
            estilo.setWrapText(true);
            estilo.setBorderBottom(BorderStyle.valueOf((short) 1));
            estilo.setBorderLeft(BorderStyle.valueOf((short) 1));
            estilo.setBorderRight(BorderStyle.valueOf((short) 1));
            estilo.setBorderTop(BorderStyle.valueOf((short) 1));

            CellStyle estiloCelda = libro.createCellStyle();
            estiloCelda.setWrapText(true);
            estiloCelda.setBorderBottom(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderLeft(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderRight(BorderStyle.valueOf((short) 1));
            estiloCelda.setBorderTop(BorderStyle.valueOf((short) 1));

            ArrayList<String[]> tablaComp = new ArrayList<>();

            for (String tipo : tipos) {
                tablaComp.addAll(DBConexion.getData(tipo));
            }

            String[] cabezera = {"ID", "Etiqueta AEMet", "Denominación", "Código Fabricante", "Cantidad", "Fecha Recepción", "Fecha Modificación", "Tipo"};

            int f = 0, c = 0;
            for (String[] fila : tablaComp) {
                Row filaX = hoja.createRow(f);
                for (String ignored : fila) {
                    Cell celdaX = filaX.createCell(c);
                    if (f == 0) {
                        filaX.setHeight((short) 450);
                        celdaX.setCellStyle(estilo);
                        celdaX.setCellValue(cabezera[c]);
                    }
                    c++;
                }
                f++;
            }

            f = 0;
            c = 0;
            int filaInicio = 1;
            for (String[] fila : tablaComp) {
                Row filaX = hoja.createRow(filaInicio);
                filaX.setHeight((short) 450);
                filaInicio++;
                for (String celda : fila) {
                    Cell celdaX = filaX.createCell(c);
                    celdaX.setCellValue(celda);
                    celdaX.setCellStyle(estiloCelda);
                    hoja.setColumnWidth(c, 4500);
                    c++;
                }
                c = 0;
                f++;
            }

            libro.write(archivo);
            archivo.close();
            Desktop.getDesktop().open(archivoXLS);
        }
    }

    public void importarExcel(JTable tabla) throws SQLException {

        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls");
        chooser.setFileFilter(filter);
        chooser.setDialogTitle("Seleccionar archivo Excel");
        int userSelection = chooser.showOpenDialog(this);
        DefaultTableModel model = (DefaultTableModel) tabla.getModel();

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();


            // Limpiar la tabla antes de la importación
            model.setRowCount(0);

            try (FileInputStream excelFile = new FileInputStream(file)) {

                HSSFWorkbook workbook = new HSSFWorkbook(excelFile);

                // Obtener la primera hoja de trabajo
                Sheet sheet = workbook.getSheetAt(0);

                // Iterar sobre las filas de la hoja de trabajo
                Iterator<Row> rowIterator = sheet.iterator();
                int it = 0;
                String[] cabezeraArr = {"ID", "Etiqueta AEMet", "Denominación", "Código Fabricante", "Cantidad", "Fecha Recepción", "Fecha Modificación", "Tipo"};
                ArrayList<String> cabezera = new ArrayList<>(List.of(cabezeraArr));

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Object[] rowData = new Object[row.getLastCellNum()];

                    // Iterar sobre las celdas de la fila
                    for (int i = 0; i < row.getLastCellNum()-1; i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        // Dependiendo del tipo de celda, obtener el valor adecuado
                        switch (cell.getCellType()) {
                            case STRING:
                                rowData[i] = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                rowData[i] = cell.getNumericCellValue();
                                break;
                            case BOOLEAN:
                                rowData[i] = cell.getBooleanCellValue();
                                break;
                            default:
                                rowData[i] = null;
                        }
                    }

                    // Verificar si alguna celda coincide con elementos de la cabecera
                    boolean coincidenciaCabecera = Arrays.stream(rowData).anyMatch(cellData -> cabezera.contains(String.valueOf(cellData)));

                    // Agregar la fila al modelo de la tabla solo si no hay coincidencia
                    if (!coincidenciaCabecera) {
                        model.addRow(rowData);
                        it++;
                    }
                }

            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error al leer el archivo Excel: \n" + ex, "Error", JOptionPane.ERROR_MESSAGE);
            }
        }

        int confirm = JOptionPane.showConfirmDialog(this, "¿Desea sobreescribir los datos actuales?");

        if (confirm != 1) {

            String tipo = pestanas.getTitleAt(pestanas.getSelectedIndex());
            DBConexion.sobreescribirTabla(tipo, model);

        }
    }

    private static Object getCellValue(Cell cell) {
        // Manejar diferentes tipos de celdas
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            default:
                return null;
        }
    }
    public static class CustomStringEditor extends DefaultCellEditor {

        public CustomStringEditor() {
            super(new JTextField());
            JTextField textField = (JTextField) getComponent();
            textField.setDocument(createDocument());
            textField.setHorizontalAlignment(SwingConstants.CENTER);
        }

        private Document createDocument() {
            PlainDocument document = new PlainDocument();

            // Use a DocumentFilter to enforce the desired format
            ((PlainDocument) document).setDocumentFilter(new DocumentFilter() {
                @Override
                public void insertString(FilterBypass fb, int offs, String str, AttributeSet a) throws BadLocationException {
                    super.insertString(fb, offs, str, a);
                }

                @Override
                public void replace(FilterBypass fb, int offset, int length, String text, AttributeSet attrs) throws BadLocationException {
                    super.replace(fb, offset, length, text, attrs);
                }
            });

            return document;
        }

        private boolean isValidString(String str) {
            // Validate the string format "ccc.nn.nnn.nn"
            String[] parts = str.split("\\.");
            if (parts.length == 4) {
                for (int i = 0; i < parts.length; i++) {
                    if (i == 0) {
                        if (!isAlpha(parts[i]) || parts[i].length() != 3) {
                            return false;
                        }
                    } else {
                        if (!isNumeric(parts[i]) || parts[i].length() != (i == 1 ? 2 : 3)) {
                            return false;
                        }
                    }
                }
                return true;
            }
            return false;
        }

        private boolean isAlpha(String str) {
            return str.matches("[a-zA-Z]+");
        }

        private boolean isNumeric(String str) {
            return str.matches("\\d+");
        }

        @Override
        public boolean stopCellEditing() {
            try {
                // Validate the final string format before stopping cell editing
                JTextField textField = (JTextField) getComponent();
                String text = textField.getText();
                if (isValidString(text)) {
                    return super.stopCellEditing();
                } else {
                    // Show a warning if the final string is not valid
                    System.out.println("bee");
                    return false;
                }
            } catch (Exception ex) {
                return false;
            }
        }
    }
}
