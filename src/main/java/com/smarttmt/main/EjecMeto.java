
package com.smarttmt.main;

import com.smarttmt.ctrl.CtrlRepo;
import com.smarttmt.interfaz.WordInte;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 *
 * @author desarrollo001
 */
public class EjecMeto {

     private static final Logger LOGGER = LogManager.getLogger(EjecMeto.class);
    
    public static void main(String[] args) {

        try {
            WordInte wordInte = new CtrlRepo();
            switch (args.length) {
                case 4:
                    wordInte.writDOT(args[0], args[1], args[2], args[3]);
                    break;
                case 1:
                    wordInte.compDire(args[0]);
                    break;
                default:
                    System.out.println("Debe enviar los paramertros de manera correcta");
                    System.out.println("Para generacion de plantillas envie 4 parametros como esta a continuacion ");
                    System.out.println("1. Parametro Path plantilla");
                    System.out.println("2. Parametro Path destino");
                    System.out.println("3. Parametro Path archivo de datos");
                    System.out.println("4. Parametro separador del archivo de datos");
                    System.out.println("");
                    System.out.println("Para comprimir directorio envie 1 parametros");
                    System.out.println("1. Parametro Path");
                    break;
            }

        } catch (Exception ex) {
            LOGGER.error("Error {}", ex.getMessage());
        }
    }

}
