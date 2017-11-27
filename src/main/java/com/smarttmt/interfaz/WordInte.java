package com.smarttmt.interfaz;

/**
 * Esta interfaz tiene un metodo que se encarga de escribir un archivo Word
 *
 * @author adalbertdavidaroca
 */
public interface WordInte {

    /**
     *
     * Se crea este metodo para realizar la escritura de un archivo en word,
     * apartir de un formato o plantilla existente en formato DOT.
     *
     * @param pathPlan ruta de la plantilla
     * @param pathDeex ruta destino donde se va almacenar resultado de la 
     *                 plantilla.
     * @param pathFida ruta del archivo de datos que se van agregar en la 
     *        plantilla.
     * @param separato separador que se usa para identificar las columnas en el
     * archivo de datos
     *
     * @throws Exception indica que el metodo genera excepciones que deben ser
     * capturadas para identificar cuando la misma no funcione bien.
     */
    public void writDOT(String pathPlan, String pathDeex, String pathFida, String separato) throws Exception;    
    
    /***
     * Metodo para realizar la compresion de un directorio que se pase por parametro.
     * @param pathDire ruta del directorio
     * @throws Exception genera la excepcion.
     */
    public void compDire(String pathDire) throws Exception;    
    
    
}
