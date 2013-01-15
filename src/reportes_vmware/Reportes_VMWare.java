/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package reportes_vmware;

import Control.Control_ppal;
import Vista.Vista_ppal;

/**
 *
 * @author SSrdespinosa
 */
public class Reportes_VMWare {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
       Vista_ppal vista = new Vista_ppal();
       Control_ppal control = new Control_ppal(vista);
       
       vista.setVisible(true);
    }
}
