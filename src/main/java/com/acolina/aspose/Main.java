package com.acolina.aspose;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;

import java.io.File;

public class Main {

    public static void main(String[] args) throws Exception {
        validFile(args);
        String type = args[0];
        if ("C".equals(type)) {
            type1(args);
        } else if ("M".equals(type)) {
            type2(args);
        } else {
            throw new UnsupportedOperationException("Tipo no valido");
        }

    }

    private static void type1(String[] args) throws Exception {
        validFile(args);
        String out2 = args[2];
        Document document = new Document(args[1]);
        document.save(out2);
    }

    private static void type2(String[] args) throws Exception {

        Document document = new Document(args[1]);
        for (int i = 2; i < args.length; i++) {
            Document appendDoc = new Document(args[i]);
            document.appendDocument(appendDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        }
        String out = args[1].substring(0, args[1].lastIndexOf(".")).concat("_merge.docx");
        document.save(out);
    }

    private static void validFile(String[] args) {
        if (args.length < 3) {
            throw new IllegalArgumentException("Argumentos no validos");
        }
        File file1 = new File(args[1]);

        if (!file1.exists()) {
            throw new RuntimeException("fichero : \'" + args[1] + "\" no existe");
        }
    }
}
