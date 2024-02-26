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

import javax.swing.*;
import javax.swing.event.TableModelEvent;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.plaf.FontUIResource;
import javax.swing.table.*;
import javax.swing.text.StyleContext;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;

public class GUI extends JFrame {
    private JTabbedPane pestanas;
    private JTextField campoBusqueda;
    private JTable tabla;
    private JTable tabla2;
    private JTable tabla3;
    private JButton botonExportar;
    private JButton botonImportar;
    private JButton botonExportarT;
    private JButton botonPrueba;
    private JPanel panelPrincipal;
    private JPanel panelBotones;
    private JPanel panelTabla;
    private JPanel panelTabla2;
    private JPanel panelTabla3;
    private JPanel panelBusqueda;

    int it = 0;

    public GUI(DBConexion con) throws SQLException {

        $$$setupUI$$$();

        String[] tipos = {"SSBB", "Informática", ""};
        ArrayList<DefaultTableModel> modeloTablas = new ArrayList<>();

        // Configuración de la ventana
        setTitle("INVENTARIO");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(1000, 1000);
        setLocationRelativeTo(null);
        setVisible(true);
        setResizable(false);

        TableRowSorter<TableModel> sorter;
        DefaultTableModel modeloTabla = new DefaultTableModel();
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

            tabla = new JTable(modeloTabla);
            tabla.setFont(new Font("", Font.PLAIN, 16));
            tabla.getTableHeader().setFont(new Font("", Font.BOLD, 16));

            sorter = new TableRowSorter<>(modeloTabla);
            tabla.setRowSorter(sorter);

            tabla.setRowHeight(25);

            GUI.CalendarioRenderer cr = new GUI.CalendarioRenderer();
            GUI.CalendarioEditor ce = new GUI.CalendarioEditor();

            tabla.getColumnModel().getColumn(5).setCellRenderer(cr);
            tabla.getColumnModel().getColumn(5).setCellEditor(ce);
            tabla.getColumnModel().getColumn(6).setCellRenderer(cr);
            tabla.getColumnModel().getColumn(6).setCellEditor(ce);
            tabla.getColumnModel().getColumn(7).setCellRenderer(new Interfaz.ButtonRenderer());
            tabla.getColumnModel().getColumn(7).setCellEditor(new Interfaz.ButtonEditor(tabla));

            it = refrescarTablas(modeloTablas, tipos, it);

            JScrollPane scrollPane = new JScrollPane(tabla);
            panelTabla.setLayout(new MigLayout("wrap, fill"));
            panelTabla.add(scrollPane, "grow");

            tabla.repaint();

            modeloTabla.addTableModelListener(e -> {

                if (e.getType() == TableModelEvent.UPDATE) {
                    int fila = e.getFirstRow();
                    int columna = e.getColumn();
                    int id = Integer.parseInt((String) tabla.getValueAt(fila, 0));

                    if (columna != 6) {

                        String nuevoValor = (String) modeloTabla.getValueAt(fila, columna);

                        con.updateTabla(id, columna, nuevoValor, tipos[0]);

                        try {
                            it = refrescarTablas(modeloTablas, tipos, it);
                        } catch (SQLException ex) {
                            throw new RuntimeException(ex);
                        }
                    }
                }
            });

        } catch (SQLException ex) {
            throw new RuntimeException(ex);
        }

        TableRowSorter<TableModel> sorter2;
        DefaultTableModel modeloTabla2 = new DefaultTableModel();
        modeloTablas.add(modeloTabla2);

        try {

            modeloTabla2.addColumn("ID");
            modeloTabla2.addColumn("Etiqueta AEMet");
            modeloTabla2.addColumn("Denominación");
            modeloTabla2.addColumn("Código Fabricante");
            modeloTabla2.addColumn("Cantidad");
            modeloTabla2.addColumn("Fecha Recepción");
            modeloTabla2.addColumn("Fecha Modificación");
            modeloTabla2.addColumn("");

            tabla2 = new JTable(modeloTabla2);
            tabla2.setFont(new Font("", Font.PLAIN, 16));
            tabla2.getTableHeader().setFont(new Font("", Font.BOLD, 16));

            sorter2 = new TableRowSorter<>(modeloTabla2);
            tabla2.setRowSorter(sorter2);

            tabla2.setRowHeight(25);

            GUI.CalendarioRenderer cr = new GUI.CalendarioRenderer();
            GUI.CalendarioEditor ce = new GUI.CalendarioEditor();

            tabla2.getColumnModel().getColumn(5).setCellRenderer(cr);
            tabla2.getColumnModel().getColumn(5).setCellEditor(ce);
            tabla2.getColumnModel().getColumn(6).setCellRenderer(cr);
            tabla2.getColumnModel().getColumn(6).setCellEditor(ce);
            tabla2.getColumnModel().getColumn(7).setCellRenderer(new Interfaz.ButtonRenderer());
            tabla2.getColumnModel().getColumn(7).setCellEditor(new Interfaz.ButtonEditor(tabla2));

            it = refrescarTablas(modeloTablas, tipos, it);

            JScrollPane scrollPane = new JScrollPane(tabla2);
            panelTabla2.setLayout(new MigLayout("wrap, fill"));
            panelTabla2.add(scrollPane, "grow");

            tabla2.repaint();

            modeloTabla2.addTableModelListener(e -> {

                if (e.getType() == TableModelEvent.UPDATE) {
                    int fila = e.getFirstRow();
                    int columna = e.getColumn();
                    int id = Integer.parseInt((String) tabla2.getValueAt(fila, 0));

                    if (columna != 6) {

                        String nuevoValor = (String) modeloTabla2.getValueAt(fila, columna);

                        con.updateTabla(id, columna, nuevoValor, tipos[1]);

                        try {
                            it = refrescarTablas(modeloTablas, tipos, it);
                        } catch (SQLException ex) {
                            throw new RuntimeException(ex);
                        }
                    }
                }
            });

        } catch (SQLException ex) {
            throw new RuntimeException(ex);
        }

        TableRowSorter<TableModel> sorter3;
        DefaultTableModel modeloTabla3 = new DefaultTableModel();
        modeloTablas.add(modeloTabla3);

        try {

            modeloTabla3.addColumn("ID");
            modeloTabla3.addColumn("Etiqueta AEMet");
            modeloTabla3.addColumn("Denominación");
            modeloTabla3.addColumn("Código Fabricante");
            modeloTabla3.addColumn("Cantidad");
            modeloTabla3.addColumn("Fecha Recepción");
            modeloTabla3.addColumn("Fecha Modificación");
            modeloTabla3.addColumn("");

            tabla3 = new JTable(modeloTabla3);
            tabla3.setFont(new Font("", Font.PLAIN, 16));
            tabla3.getTableHeader().setFont(new Font("", Font.BOLD, 16));

            sorter3 = new TableRowSorter<>(modeloTabla3);
            tabla3.setRowSorter(sorter3);

            tabla3.setRowHeight(25);

            GUI.CalendarioRenderer cr = new GUI.CalendarioRenderer();
            GUI.CalendarioEditor ce = new GUI.CalendarioEditor();

            tabla3.getColumnModel().getColumn(5).setCellRenderer(cr);
            tabla3.getColumnModel().getColumn(5).setCellEditor(ce);
            tabla3.getColumnModel().getColumn(6).setCellRenderer(cr);
            tabla3.getColumnModel().getColumn(6).setCellEditor(ce);
            tabla3.getColumnModel().getColumn(7).setCellRenderer(new Interfaz.ButtonRenderer());
            tabla3.getColumnModel().getColumn(7).setCellEditor(new Interfaz.ButtonEditor(tabla3));

            it = refrescarTablas(modeloTablas, tipos, it);

            JScrollPane scrollPane = new JScrollPane(tabla3);
            panelTabla3.setLayout(new MigLayout("wrap, fill"));
            panelTabla3.add(scrollPane, "grow");

            tabla3.repaint();

            modeloTabla3.addTableModelListener(e -> {

                if (e.getType() == TableModelEvent.UPDATE) {
                    int fila = e.getFirstRow();
                    int columna = e.getColumn();
                    int id = Integer.parseInt((String) tabla3.getValueAt(fila, 0));

                    if (columna != 6) {

                        String nuevoValor = (String) modeloTabla3.getValueAt(fila, columna);

                        con.updateTabla(id, columna, nuevoValor, tipos[2]);

                        try {
                            it = refrescarTablas(modeloTablas, tipos, it);
                        } catch (SQLException ex) {
                            throw new RuntimeException(ex);
                        }
                    }
                }

            });

        } catch (SQLException ex) {
            throw new RuntimeException(ex);
        }

        campoBusqueda.addActionListener(e -> {

            RowFilter<TableModel, Integer> rowFilter = RowFilter.regexFilter(campoBusqueda.getText());

            switch (pestanas.getSelectedIndex()) {

                case 0:
                    sorter.setRowFilter(rowFilter);
                    break;

                case 1:
                    sorter2.setRowFilter(rowFilter);
                    break;

                case 2:
                    sorter3.setRowFilter(rowFilter);
                    break;
            }

        });

        // Botón que exporta a Excel
        botonExportar.setFont(new Font("Arial", Font.PLAIN, 16));
        botonExportar.setPreferredSize(new Dimension(150, 50));


        botonExportar.addActionListener(e -> {

            try {

                switch (pestanas.getSelectedIndex()) {

                    case 0:
                        exportarExcel(tabla);
                        break;

                    case 1:
                        exportarExcel(tabla2);
                        break;

                    case 2:
                        exportarExcel(tabla3);
                        break;
                }

            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }

        });

        // Botón que importa a Excel
        botonImportar.setFont(new Font("Arial", Font.PLAIN, 16));
        botonImportar.setPreferredSize(new Dimension(150, 50));
//        panelBotones.add(botonImportar);

        botonImportar.addActionListener(e -> {

            try {
                switch (pestanas.getSelectedIndex()) {

                    case 0:
                        importarExcel(tabla);
                        break;

                    case 1:
                        importarExcel(tabla2);
                        break;

                    case 2:
                        importarExcel(tabla3);
                        break;
                }
            } catch (SQLException ex) {
                throw new RuntimeException(ex);
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

        // Agregar los paneles al panel principal
//        panelPrincipal.add(panelBotones, "wrap");
//        panelPrincipal.add(panelBusqueda, "wrap");
        pestanas.addTab(tipos[0], panelTabla);
        pestanas.addTab(tipos[1], panelTabla2);
        pestanas.addTab(tipos[2], panelTabla3);

        // Agregar el panel principal a la ventana
        add(panelPrincipal);

        // Una última vez para que queden bien antes de poder interactuar, después se realizará
        // de forma automática
        refrescarTablas(modeloTablas, tipos, it);

        // Hacer visible la ventana
        setVisible(true);
    }

    private int refrescarTablas(ArrayList<DefaultTableModel> modeloTablas, String[] tipos, int it) throws SQLException {

        int i = 0;

        for (DefaultTableModel modeloTabla : modeloTablas) {

            it = refrescarTabla(modeloTabla, tipos[i], it);
            i++;
        }

        return it;
    }

    private int refrescarTabla(DefaultTableModel modeloTabla, String tipo, int it) throws SQLException {

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

        }

        // Esto hace falta porque el addRow requiere un array aunque sea vacío
        ArrayList<String[]> arrayVacio = new ArrayList<>();
        String[] vacio = {it + 1 + ""};
        arrayVacio.add(vacio);
        modeloTabla.addRow(arrayVacio.get(0));

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
        panelTabla = new JPanel();
        panelTabla.setLayout(new GridLayoutManager(1, 1, new Insets(0, 0, 0, 0), -1, -1));
        pestanas.addTab("Untitled", panelTabla);
        panelTabla2 = new JPanel();
        panelTabla2.setLayout(new GridLayoutManager(1, 1, new Insets(0, 0, 0, 0), -1, -1));
        pestanas.addTab("Untitled", panelTabla2);
        panelTabla3 = new JPanel();
        panelTabla3.setLayout(new GridLayoutManager(1, 1, new Insets(0, 0, 0, 0), -1, -1));
        pestanas.addTab("Untitled", panelTabla3);
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
        botonPrueba.setText("Plantilla");
        panelBotones.add(botonPrueba, new GridConstraints(1, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
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
            button.addActionListener(e -> fireEditingStopped());
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

    public static class DateLabelFormatter extends JFormattedTextField.AbstractFormatter {

        private final String datePattern = "dd-MM-yyyy";
        private final SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);

        @Override
        public Object stringToValue(String text) throws ParseException {
            return dateFormatter.parseObject(text);
        }

        @Override
        public String valueToString(Object value) {
            if (value != null) {
                Calendar cal = (Calendar) value;
                return dateFormatter.format(cal.getTime());
            }

            return "";
        }

    }

    @SuppressWarnings("ResultOfMethodCallIgnored")
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
                    } else {
                        celda.setCellValue(String.valueOf(tabla.getValueAt(f, c)));
                    }
                }
            }

            libro.write(archivo);
            archivo.close();
            Desktop.getDesktop().open(archivoXLS);

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

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();

            DefaultTableModel model = (DefaultTableModel) tabla.getModel();

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
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        if (it == 0 && !cabezera.contains(rowData[i].toString())) throw new IOException();

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

                    // Agregar la fila al modelo de la tabla
                    model.addRow(rowData);
                    it++;
                }

            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error al leer el archivo Excel: \n" + ex, "Error", JOptionPane.ERROR_MESSAGE);
            }
        }

        int confirm = JOptionPane.showConfirmDialog(this, "¿Desea sobreescribir los datos actuales?");

        if (confirm != 1) {

            String tipo = pestanas.getTitleAt(pestanas.getSelectedIndex());
            it = 0;
            DBConexion.sobreescribirTabla(tipo, (DefaultTableModel) tabla.getModel());

        }

    }
}
