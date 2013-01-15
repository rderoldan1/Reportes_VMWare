/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package Control;

import Vista.Vista_ppal;
import com.vmware.vim25.mo.ServiceInstance;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.net.URL;

/**
 *
 * @author SSrdespinosa
 */
public class Control_ppal {
    private Vista_ppal m_vista;

    public Control_ppal(Vista_ppal vista ) {
        m_vista = vista;        
        vista.addSubmitListener(new SubmitListener());
    }
    
    class SubmitListener implements ActionListener{
        public void actionPerformed(ActionEvent e){
            String user = "";
            String pass = "";
            String server = "";
            System.out.println(pass);
            try{
              user = m_vista.getUser();  
              pass = m_vista.getPassword();
              server = m_vista.getDNSName();
              m_vista.setLogText(user+pass+server); 
              loggin(user, pass, server); 
            }catch(Exception ex){
                
            }
        }
    }
    
    public void loggin(String user, String pass, String ip){
                
        ServiceInstance si = null;
        //log("Ingresando a "+ ip);
        try{
            log("Ingresando a "+ ip);
            si = new ServiceInstance(new URL("https://"+ip+"/sdk"), user, pass, true);
            //libro(si);
            si.getServerConnection().logout();
            log("Reporte generado exitosamente");
            progress(0);
        }catch(Exception e){
            log("Los datos de ingreso parecen ser erroneos");
        }
        
    }
    
    private void progress(int width){
        m_vista.setPercetCompleted(width);
    }
    
    private void log(String message){
        m_vista.setAppendLogText(message+"\n");
    }
    
}
