package de.nilswitt.oks.spltXlsxComments;

import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
        System.out.println("Start");

        XlsxHandler xlsxHandler = new XlsxHandler("./data/in.xlsx");

        xlsxHandler.read();
        xlsxHandler.splitToRecipients();
        xlsxHandler.createRecipientFiles();
    }
}
