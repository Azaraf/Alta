/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ahba;

import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

/**
 *
 * @author azaraf
 */
public class CrearExcel {
    private File file;
    private String nombre;
    private Object[][] data;
    
    public CrearExcel(String nombre, Object[][] data) {
        showSaveDialog();
    }
    
    
    private void showSaveDialog(){
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogType(JFileChooser.SAVE_DIALOG);
        fileChooser.setSelectedFile(new File("/" + nombre + new SimpleDateFormat("ddMMyyyy").format(new Date()) +".csv"));
        
        if(fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION){
            file = fileChooser.getSelectedFile();
            crearExcel();
        }
    }
    
    private void crearExcel(){
        try {
            file.createNewFile();
            try (FileOutputStream fout = new FileOutputStream(file)) {
                for (Object[] lst : data) {
                    for (Object currentValue : lst) {
                        fout.write((!currentValue.equals(lst[lst.length - 1]) ? currentValue.toString() + "," : currentValue.toString()).getBytes());
                    }
                    fout.write(("\n").getBytes());
                }
                fout.flush();
            }
        }catch(IOException e){
            e.printStackTrace();
        }
        Object[] options = {"Abrir",
            "No"};
        int open = JOptionPane.showOptionDialog(
                null,
                "¿Desea abrir el reporte?",
                "Reporte guardado",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                options,
                options[0]);
        if (open == JOptionPane.YES_OPTION) {
            try {
                Desktop.getDesktop().open(file);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Error en el archivo");
            }
        }
    }
    
    
}
