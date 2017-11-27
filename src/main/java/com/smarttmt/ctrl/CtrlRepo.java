/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.smarttmt.ctrl;

import com.smarttmt.interfaz.WordInte;
import com.smarttmt.utils.CompDire;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.zip.ZipOutputStream;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 *
 * @author desarrollo001
 */
public class CtrlRepo implements WordInte {

    private static final Logger LOGGER = LogManager.getLogger(CtrlRepo.class);

    /**
     * Se implementa el metodo writeDOT de la interfaz.
     *
     * @param pathPlan ruta de la plantilla
     * @param pathDeex ruta destino donde se va almacenar resultado de la
     * plantilla.
     * @param pathFida ruta del archivo de datos que se van agregar en la
     * plantilla.
     * @param separato separador que se usa para identificar las columnas en el
     * archivo de datos
     *
     * @throws Exception indica que el metodo genera excepciones que deben ser
     * capturadas para identificar cuando la misma no funcione bien.
     */
    @Override
    public void writDOT(String pathPlan, String pathDeex, String pathFida, String separato) throws Exception {

        LOGGER.info("Iniciando escritura del uso de la plantilla DOT para generar un archivo Word");
        LOGGER.debug("Ruta de la plantilla {}", pathPlan);
        LOGGER.debug("Ruta donde se va a escribir la plantilla {} ", pathDeex);
        LOGGER.debug("Ruta de donde se van a obtener los datos {} ", pathFida);
        LOGGER.debug("Separador para identificar columnas por cada registro del archivo de datos {} ", separato);

        /**
         * La variable archOrig se usa para verificar si la plantilla existe
         * antes de iniciar su generacion
         */
        File archOrig = null;
        /**
         * La variable archDest se usa para validar si la ruta donde se va a
         * depositar el word creado.
         */
        File archDest = null;

        /**
         * La variable pathDefi se usa para almacenar la ruta del directorio que
         * se obtiene del parametro pathDeex.
         */
        String pathDefi = null;

        /**
         * La variable fs carga un stream del archivo o plantilla para
         * manipularlo.
         */
        POIFSFileSystem fs = null;

        /**
         * La variable planWord se utiliza para manipular la plantilla y hacer
         * el reemplazo de los valores.
         */
        HWPFDocument planWord = null;

        Range range = null;

        //set de caracteres
        Charset charset = null;

        //lineas del archivo
        List<String> lines = null;

        OutputStream out = null;

        /**
         * si se cambia el set de caracteres pasa a true. Sino false
         */
        boolean chanseca = false;

        try {

            archOrig = new File(pathPlan);

            if (!archOrig.exists()) {
                LOGGER.debug("Plantilla no existe en la ruta {} ", pathPlan);
                throw new IOException("La plantilla no existe en la ruta " + pathPlan);
            }
            LOGGER.debug("Se va  obtener el path de la ruta {}. Separado {}  ", pathDeex, File.separator);
            pathDefi = pathDeex.substring(0, pathDeex.lastIndexOf(File.separator));
            LOGGER.debug("el path es {} ", pathDefi);

            archDest = new File(pathDefi);

            if (!archDest.exists()) {
                LOGGER.debug("Ruta  {} no existe donde se va a depositar el WORD , se va a crear", pathDefi);
                archDest.mkdirs();
            }

            fs = new POIFSFileSystem(archOrig);

            planWord = new HWPFDocument(fs);

            range = planWord.getRange();

            try {
                lines = Files.readAllLines(Paths.get(pathFida));
                LOGGER.debug("Se cargaron del archivo un total de {} filas", lines.size());
            } catch (IOException ex) {
                LOGGER.warn("El set de caracteres por defecto no funciono {}", ex.getMessage());
                chanseca = true;
            }

            if (chanseca) {
                charset = Charset.forName("ISO-8859-1");
                LOGGER.debug("se cambio el set de caracteres {}", charset.displayName());
                lines = Files.readAllLines(Paths.get(pathFida), charset);
                LOGGER.debug("Se cargaron del archivo un total de {} filas", lines.size());
            }

            LOGGER.debug("Empieza a recorrer lineas");

            for (String line : lines) {
                LOGGER.debug("Linea {}", line);
                if (line.indexOf(separato) > 0) {
                    String[] parts = line.split("[" + separato + "]");
                    LOGGER.debug("numero de campos por linea {} ", parts.length);
                    if (parts.length > 0) {
                        LOGGER.debug("Linea; label {} value {} ", parts[0], parts[1]);
                        planWord.getHeaderStoryRange().replaceText("[" + parts[0] + "]", parts[1]);
                        planWord.getFootnoteRange().replaceText("[" + parts[0] + "]", parts[1]);
                        range.replaceText("[" + parts[0] + "]", parts[1]);
                    }
                }
                LOGGER.debug("Termino linea");

            }

            LOGGER.info("Ruta del archivo a crear {}", pathDeex);
            archDest = new File(pathDeex);

            if (!archDest.exists()) {
                LOGGER.info("No existe el archivo en la ruta  {}, se procede a la creacion ", pathDeex);
                archDest.createNewFile();
            } else {

                LOGGER.info("el archivo que se requiere crear, ya existe {} se va a recrear", pathDeex);
                archDest.delete();
                LOGGER.info("archivo en la ruta {}, borrado", pathDeex);
                archDest.createNewFile();
                LOGGER.info("archivo en la ruta {}, se vuelve a crear", pathDeex);

            }

            out = new FileOutputStream(new File(pathDeex));
            planWord.write(out);

        } catch (Exception e) {
            LOGGER.error("Error en la generacion {} ", e.getMessage());
            throw new Exception("Error en la generacion de la plantilla [" + e.getMessage() + "]");
        } finally {
            if (planWord != null) {
                planWord.close();
            }
            if (fs != null) {
                fs.close();
            }

            if (out != null) {
                out.flush();
                out.close();
            }
        }
    }

    @Override
    public void compDire(String pathDire) throws Exception {

        LOGGER.info("Ejecutando compresion de archivo");
        Path sourceDir;
        ZipOutputStream zos;
        
        try {
            sourceDir = Paths.get(pathDire);
            LOGGER.info("Path de directorio a comprimir {} ",pathDire);
            String zipFileName = pathDire.concat(".zip");
            LOGGER.info("Nombre del archivo {} ",zipFileName);
            zos = new ZipOutputStream(new FileOutputStream(zipFileName));

            Files.walkFileTree(sourceDir, new CompDire(sourceDir,zos));

            zos.close();
        } catch (IOException ex) {
            LOGGER.error("Error en la compresion {} ", ex.getMessage());
        }
    }

}
