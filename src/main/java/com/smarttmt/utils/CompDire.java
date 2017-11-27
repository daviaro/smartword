/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.smarttmt.utils;

import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 *
 * @author desarrollo001
 */
public class CompDire extends SimpleFileVisitor<Path> {

    private final ZipOutputStream zos;

    private final Path sourceDir;

    public CompDire(Path sourceDir, ZipOutputStream zos) {
        this.sourceDir = sourceDir;
        this.zos = zos;
    }

    @Override
    public FileVisitResult visitFile(Path file,
            BasicFileAttributes attributes) {

        Path targetFile;
        
        try {
            
            targetFile = sourceDir.relativize(file);

            zos.putNextEntry(new ZipEntry(targetFile.toString()));

            byte[] bytes = Files.readAllBytes(file);
            zos.write(bytes, 0, bytes.length);
            zos.closeEntry();

        } catch (IOException ex) {
            System.err.println(ex);
        }

        return FileVisitResult.CONTINUE;
    }

}
