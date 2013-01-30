/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package Control;

import Vista.Vista_ppal;
import com.vmware.vim25.InvalidProperty;
import com.vmware.vim25.mo.Datacenter;
import com.vmware.vim25.mo.Folder;
import com.vmware.vim25.mo.HostSystem;
import com.vmware.vim25.mo.InventoryNavigator;
import com.vmware.vim25.mo.ManagedEntity;
import com.vmware.vim25.mo.ServiceInstance;
import com.vmware.vim25.mo.VirtualMachine;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.Calendar;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author SSrdespinosa
 */
public class Control_ppal {
    private Vista_ppal m_vista;
     String ip = "";

    public Control_ppal(Vista_ppal vista ) {
        m_vista = vista;        
        vista.addSubmitListener(new SubmitListener());
        
    }
    
    class SubmitListener implements ActionListener{
        public void actionPerformed(ActionEvent e){
            String user = "";
            String pass = "";
            ip = "";
            //System.out.println(pass);
            try{
              user = m_vista.getUser();  
              pass = m_vista.getPassword();
              ip = m_vista.getDNSName();
              //m_vista.setLogText(user+pass+ip); 
              loggin(user, pass, ip); 
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
            libro(si);
            si.getServerConnection().logout();
            log("Reporte generado exitosamente");
            progress(0);
        }catch(Exception e){
            log("Los datos de ingreso parecen ser erroneos");
        }
        
    }
    
    public void libro(ServiceInstance si) throws IOException{
     
        try {
            Workbook wb = new HSSFWorkbook();
            Sheet sheet1 = wb.createSheet("Recursos Asignados");
            Sheet sheet2 = wb.createSheet("Balance Memoria Procesador");
            Sheet sheet3 = wb.createSheet("Almacenamiento");
            sheet1 = VCenterInfo(si, sheet1);
            progress(30 - m_vista.getPercetCompleted());
            sheet2 = hoja2(si, sheet2);
            progress(60 - m_vista.getPercetCompleted());
            sheet3 = hoja3(si, sheet3);
            progress(90 - m_vista.getPercetCompleted());
            autoSize(wb);
            
            Calendar cal = Calendar.getInstance();
            int month = cal.get(Calendar.MONTH) + 1;
            int year = cal.get(Calendar.YEAR);
            FileOutputStream fileOut;
            progress(95 - m_vista.getPercetCompleted());
            String a = System.getProperty("user.home")+"\\Desktop";
            createDirectoryIfNeeded(a+"\\Reportes");
            System.out.println(a);
            String name = a+"\\Reportes\\"+year+"\\"+month+"\\"+createName()+".xls";
            fileOut = new FileOutputStream(name);
            wb.write(fileOut);
            fileOut.close();
            progress(100 - m_vista.getPercetCompleted());
            log("Reporte creado exitosamente:"+ name);
        } catch (FileNotFoundException ex) {
            
        }
        progress(0);
    }
    
    
    /**
     * Metodo para crear hoja 1 del reporte, resumen de plataforma virtual
     * @param si
     * @param sheet1
     * @return
     * @throws InvalidProperty 
     */
    public Sheet VCenterInfo(ServiceInstance si, Sheet sheet1) throws InvalidProperty{
        try{
            log("Creando la hoja 1");
            Folder rootFolder = si.getRootFolder();
            ManagedEntity [] dataCenters = si.getRootFolder().getChildEntity();
            
       
            int inicial = 2;
            Row row = null;
            Cell cell = null;
            
            
            row = sheet1.createRow(1);
            row.createCell(0).setCellValue("Server");
            row.createCell(1).setCellValue("Estado");
            row.createCell(2).setCellValue("Cluster");
            row.createCell(3).setCellValue("Tipo");
            row.createCell(4).setCellValue("Procesador Asignado (Cores)");
            row.createCell(5).setCellValue("Memoria Asignado (GB)");
            row.createCell(6).setCellValue("Maquinas Virtuales");
            row.createCell(7).setCellValue("Modelo");
            
            int num_hosts = 0;
            int mem = 0;
            int proc = 0;
            int num_vms = 0;
            for(int i = 0; i < dataCenters.length; i++){
                Datacenter dataCenter = (Datacenter)dataCenters[i];
                
                
                ManagedEntity [] hosts = new InventoryNavigator(dataCenter)
                        .searchManagedEntities("HostSystem");
                
                for(int j = 0; j < hosts.length; j++){
                    row = sheet1.createRow(inicial);
                    HostSystem host = (HostSystem) hosts[j];
                    int  num_machines = host.getVms().length;
                    num_vms += num_machines;
                    ManagedEntity [] vms = new InventoryNavigator(host)
                            .searchManagedEntities("VirtualMachine");
                    int cores = 0;
                    int ram = 0;
                    for(int sa = 0; sa < vms.length; sa++){
                        VirtualMachine virtualMachine = (VirtualMachine)
                                vms[sa];
                        ram += virtualMachine.getConfig()
                                .getHardware().getMemoryMB()/1024;
                        cores += virtualMachine.getConfig()
                                .getHardware().getNumCPU();
                    }
                    mem += ram;
                    proc += cores;
                    
                    
                    row.createCell(0).setCellValue(host.getName());
                    row.createCell(1).setCellValue(host.getRuntime()
                            .getPowerState().toString());
                    row.createCell(2).setCellValue(dataCenter.getName());                    
                    row.createCell(3).setCellValue(host.getParent().getName());
                    row.createCell(4).setCellValue(cores);
                    row.createCell(5).setCellValue(ram);                    
                    row.createCell(6).setCellValue(num_machines);
                    row.createCell(7).setCellValue(host.getSummary()
                            .getHardware().getModel());
                    
                    inicial++;
                    System.out.println(""+inicial);
                    num_hosts += 1;
                }
                
                
                
                
            }
            log(num_hosts + " Hosts");
            
            row = sheet1.createRow(inicial);
            row.createCell(0).setCellValue(num_hosts+" Hosts");
            row.createCell(4).setCellValue(proc);
            row.createCell(5).setCellValue(mem);
            row.createCell(6).setCellValue(num_vms);
            inicial++;
            row = sheet1.createRow(inicial);
            row.createCell(0).setCellValue("Maquina Virtual");
            row.createCell(1).setCellValue("Estado");
            row.createCell(2).setCellValue("IP");
            row.createCell(3).setCellValue("Host");
            row.createCell(4).setCellValue("Memoria Asignada (GB)");
            row.createCell(5).setCellValue("Procesadores Asigandos (Cores)");
            row.createCell(6).setCellValue("OS");
            inicial++;
            int vm = 0;
            
            for(int i = 0; i < dataCenters.length; i++){
                Datacenter dataCenter = (Datacenter)dataCenters[i];
                ManagedEntity [] hosts = new InventoryNavigator(dataCenter)
                        .searchManagedEntities("HostSystem");
               
                
                for(int k = 0; k < hosts.length; k++){
                    progress(1);
                    ManagedEntity [] virtualMachines = new
                            InventoryNavigator(hosts[k])
                            .searchManagedEntities("VirtualMachine");
                    
                    vm += virtualMachines.length;
                    for(int j = 0; j < virtualMachines.length; j++){
                        row = sheet1.createRow(inicial);
                        VirtualMachine virtualMachine = (VirtualMachine)
                                virtualMachines[j];
                        row.createCell(0).setCellValue(virtualMachine.getName());
                        row.createCell(1).setCellValue(virtualMachine.getRuntime()
                                .getPowerState().toString());
                        row.createCell(2).setCellValue(virtualMachine.getGuest()
                                .getIpAddress());
                        row.createCell(3).setCellValue(hosts[k].getName());
                        row.createCell(4).setCellValue(virtualMachine.getConfig()
                                .getHardware().getMemoryMB()/1024);
                        row.createCell(5).setCellValue(virtualMachine.getConfig()
                                .getHardware().getNumCPU());
                        row.createCell(6).setCellValue(virtualMachine.getGuest()
                                .getGuestFullName());                        
                        inicial++;
                    }
                }
            }
            row = sheet1.createRow(inicial);
            row.createCell(0).setCellValue(vm+" Virtual Machines");
            log(vm+" Maquinas Virtuales");
     
     
            
        }catch(Exception e){
            System.out.print(e);
        }
       return sheet1; 
    }
    
    /**
     * Metodo para crear la hoja 2 del reporte, calculos de consumo en host
     * @param si
     * @param sheet
     * @return 
     */
    public Sheet hoja2 (ServiceInstance si, Sheet sheet) {
        try {
            log("Creando la hoja 2");
            Row row = null;
            Cell cell = null;
            
            row = sheet.createRow(0); 
            row.createCell(7).setCellValue("% HA");
            row.createCell(8).setCellValue(20);
            
            row = sheet.createRow(1); 
            row.createCell(0).setCellValue("Nombre");
            row.createCell(1).setCellValue("Base (Fisica)");
            sheet.addMergedRegion(new CellRangeAddress(1,1,1,2));
            row.createCell(3).setCellValue("Base (Fisica Real)");
            sheet.addMergedRegion(new CellRangeAddress(1,1,3,4));
            row.createCell(5).setCellValue("Actual (Uso)");
            sheet.addMergedRegion(new CellRangeAddress(1,1,5,6));
            row.createCell(7).setCellValue("Disponible con HA");
            sheet.addMergedRegion(new CellRangeAddress(1,1,7,8));
           
            
            row = sheet.createRow(2); 
            row.createCell(1).setCellValue("Procesador (core)");
            row.createCell(2).setCellValue("Memoria GB");
            row.createCell(3).setCellValue("Procesador (core)");
            row.createCell(4).setCellValue("Memoria GB");
            row.createCell(5).setCellValue("Procesador %");
            row.createCell(6).setCellValue("Memoria %");
            row.createCell(7).setCellValue("Procesador");
            row.createCell(8).setCellValue("Memoria");
            row.createCell(9).setCellValue("Core Asignados");
            row.createCell(10).setCellValue("Maquinas Virtuales");
            row.createCell(11).setCellValue("Promedio cores por maquina");
            
            int inicial = 3;
            Folder rootFolder = si.getRootFolder();
            ManagedEntity [] hosts = new InventoryNavigator(rootFolder)
                            .searchManagedEntities("HostSystem");
            for(int i = 0; i < hosts.length; i++){
                HostSystem host = (HostSystem) hosts[i];
                
                log(host.getName());
                progress(1);
                row = sheet.createRow(inicial);
                row.createCell(0).setCellValue(host.getName());
                row.createCell(1).setCellValue(host.getHardware().getCpuInfo().getNumCpuCores());
                row.createCell(2).setCellValue(Integer.parseInt(readableFileSize(host.getHardware().getMemorySize())));
               // row.createCell(3).setCellValue(host.getHardware().getCpuInfo().getNumCpuCores()* 0.75);
               // row.createCell(4).setCellValue(Integer.parseInt((readableFileSize(host.getHardware().getMemorySize())))*0.75);
                row.createCell(3).setCellFormula("(100-I1)%*B"+(inicial+1));
                row.createCell(4).setCellFormula("(100-I1)%*C"+(inicial+1));
                
                double mem_usage = (double) host.getSummary().getQuickStats().getOverallMemoryUsage()/1024;
                double mem_total = (double) host.getSummary().getHardware().getMemorySize()/1073741824;
                double cpu_usage = (double) host.getSummary().getQuickStats().getOverallCpuUsage();
                double cpu_total = (double) host.getSummary().getHardware().getCpuMhz();
                double a = cpu_usage/cpu_total;
                int cores = host.getHardware().getCpuInfo().getNumCpuCores();
                
                row.createCell(5).setCellValue((cpu_usage/(cpu_total*cores))*100);
                row.createCell(6).setCellValue((mem_usage/mem_total)*100);
                
                row.createCell(7).setCellFormula("(100-F"+(inicial+1)+")%*D"+(inicial +1));
                row.createCell(8).setCellFormula("(100-G"+(inicial+1)+")%*E"+(inicial +1));
                
                int cores_vm = 0;
                ManagedEntity [] virtualMachines = new
                            InventoryNavigator(host)
                            .searchManagedEntities("VirtualMachine");
                  for(int j = 0; j < virtualMachines.length; j++){
                      VirtualMachine vm = (VirtualMachine) virtualMachines[j];
                      cores_vm += vm.getConfig().getHardware().getNumCPU();
                  }
                  
                row.createCell(9).setCellValue(cores_vm);  
                row.createCell(10).setCellValue(virtualMachines.length); 
                row.createCell(11).setCellFormula("J"+(inicial+1)+"/K"+(inicial+1));
                
                inicial++;
            }
            
            
        } catch (Exception e) {
            log(e.toString());
        }
        return sheet;
    }
    /**
     * Metodo que crea la hoja 3 del reporte, esta solo contiene formulas
     * @param si
     * @param sheet
     * @return 
     */
     public Sheet hoja3 (ServiceInstance si, Sheet sheet){
         try{
            log("Creando la hoja 3");
            Row row = null;
            Cell cell = null;
            
            row = sheet.createRow(0); 
             sheet.addMergedRegion(new CellRangeAddress(0,0,0,6));
            //row.createCell(0).setCellValue("Información EVA Servicios Nutresa");
            
            row = sheet.createRow(1);
            row.createCell(0).setCellValue("% HA");
            row.createCell(1).setCellValue(20);
            sheet.addMergedRegion(new CellRangeAddress(1,1,2,6));
            row.createCell(2).setCellValue("FiberChannel");
            
            row = sheet.createRow(2);
            row.createCell(0).setCellValue("Dispositivo");
            row.createCell(1).setCellValue("Número de Discos");
            row.createCell(2).setCellValue("Tamaño de Cada Disco");
            row.createCell(3).setCellValue("Capacidad Total");
            row.createCell(4).setCellValue("Tamaño del Arreglo");
            row.createCell(5).setCellValue("Capacidad Utilizable");
            row.createCell(6).setCellValue("Capacidad Efectiva");
            row.createCell(7).setCellValue("Capacidad Utilizada");
            row.createCell(8).setCellValue("% Capacidad Utilizada");
            row.createCell(9).setCellValue("Capacidad Disponible");
            row.createCell(10).setCellValue("% Capacidad Disponible");
            
            log("Creando fórmulas");
            row = sheet.createRow(3);
            row.createCell(0).setCellValue("Ejemplo: EVA 1");
            row.createCell(1).setCellValue("Ingrese este valor");
            row.createCell(2).setCellValue("Ingrese este valor");
            row.createCell(3).setCellFormula("B4*C4");
            row.createCell(4).setCellFormula("((D4)*(((B4-1)/B4)*10000)/100/100)");
            row.createCell(5).setCellFormula("((D4*1000000000)/1073741824*(((B4-1)/B4)*10000)/100/100)");
            row.createCell(6).setCellFormula("F4-(F4*D2%)");
            row.createCell(7).setCellValue("Ingrese este valor");
            row.createCell(8).setCellFormula("H4*100/G4");
            row.createCell(9).setCellFormula("G4-H4");
            row.createCell(10).setCellFormula("J4*100/G4");
            
            log("Creando fórmulas");
            row = sheet.createRow(4);
            row.createCell(0).setCellValue("Ejemplo: P2000");
             row.createCell(1).setCellValue("Ingrese este valor");
            row.createCell(2).setCellValue("Ingrese este valor");
            row.createCell(3).setCellFormula("B5*C5");
            row.createCell(4).setCellFormula("((D5)*(((B5-1)/B5)*10000)/100/100)");
            row.createCell(5).setCellFormula("((D5*1000000000)/1073741824*(((B5-1)/B5)*10000)/100/100)");
            row.createCell(6).setCellFormula("F5-(F5*D2%)");
            row.createCell(7).setCellValue("Ingrese este valor");
            row.createCell(8).setCellFormula("H5*100/G5");
            row.createCell(9).setCellFormula("G5-H5");
            row.createCell(10).setCellFormula("J5*100/G5");
            
         }catch(Exception e){
             
         }
         return sheet;
    }
    
    private void progress(int increment){
        int actual = m_vista.getPercetCompleted();
        m_vista.setPercetCompleted(actual + increment);
        
    }
    
    private void log(String message){
        m_vista.setAppendLogText(message+"\n");
    }
    
    private void createDirectoryIfNeeded(String directoryName){
         Calendar cal = Calendar.getInstance();
         int month = cal.get(Calendar.MONTH) + 1;
         int year = cal.get(Calendar.YEAR);
         
         File theDir = new File(directoryName);
         File dirYear = new File(directoryName+"\\"+year);
         File dirMonth = new File(directoryName+"\\"+year+"\\"+month);
         
         // if the directory does not exist, create it
         if (!theDir.exists())
         {
            
             log("Creando carpeta para Reportes");
             theDir.mkdir();
         }
         if (!dirYear.exists())
         {
             
             log("Creando carpeta para año " );
             dirYear.mkdir();
         }
         if (!dirMonth.exists())
         {
            
             log("Creando carpeta para mes ");
             dirMonth.mkdir();
         }
     }
    
     /**
      * Metodo para crear el nombre del archivo 
      * @return nombre
      */
     public String createName(){
        String server;
        Calendar cal = Calendar.getInstance();
        int day = cal.get(Calendar.DATE);
        int month = cal.get(Calendar.MONTH) + 1;
        int year = cal.get(Calendar.YEAR);
        int hour = cal.get(Calendar.HOUR_OF_DAY);
        int minute = cal.get(Calendar.MINUTE);
         
        server = ip+"_"+year+"_"+month+"_"+day+"_"+hour+minute;
         
         return server;
     }
     
      /**
      * Metodo para convertir un numero de tipo Long a un entero con las 
      * unidades en bytes
      * @param size
      * @return 
      */
     public String readableFileSize(long size) {
         if(size <= 0) return "0";
         final String[] units = new String[] { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
         int digitGroups = (int) (Math.log10(size)/Math.log10(1024));
         String result = null;
         result = new DecimalFormat("#,##0.#").format(size/Math.pow(1024, digitGroups));
         return result;
     }
     
          /**
      * Metodo para definir el ancho de las columnas de la hoja de excel se 
      * ajusten al contenido
      * @param workbook 
      */
     public void autoSize(Workbook workbook){
         for(int i = 0; i <= 2; i++){
             Sheet sheet = workbook.getSheetAt(i);
             for(int j = 0; j< 20; j++){
                 sheet.autoSizeColumn(j);
             }
             
         }
     }
    
}
