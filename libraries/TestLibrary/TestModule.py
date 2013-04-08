# -*- coding: utf-8 -*-
import uno

def getText():
    return "Hello World !"

def HelloWorldWriter( ):
    """Prints the string 'Hello World(in Python)' into the current document"""

    # get the doc from the scripting context which is made available to all scripts
    model = XSCRIPTCONTEXT.getDocument()
    
    # Get current document with UNO service
    # ctx = uno.getComponentContext()
    # smgr = ctx.ServiceManager
    # desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    # model = desktop.getCurrentComponent()

    # Set Text to the cursor Writer
    text = model.Text
    cursor = text.createTextCursor()
    text.insertString( cursor, getText(), 0 )


def HelloWorldCalc(evt):
    # Get document : <Interface : xSpreadsheetDocument>
    xSpreadsheetDocument = XSCRIPTCONTEXT.getDocument()
    # Get the collection of sheets in the document : <Interface : xSpreadsheets>
    xSpreadsheets = xSpreadsheetDocument.getSheets()
    # Get first sheet
    # xIndexAccess = xSpreadsheets.createInstance("com.sun.star.XIndexAccess")
    # xSpreadsheet = xIndexAccess.getByIndex(0)
    xSpreadsheet = xSpreadsheets.getByIndex(0)

    xCell = xSpreadsheet.getCellByPosition(1, 1)
    xCell.setValue(100)


    # System.out.println("Getting spreadsheet") ;

    # XSpreadsheets xSheets = myDoc.getSheets() ;
    # XIndexAccess oIndexSheets = (XIndexAccess) UnoRuntime.queryInterface(
    #     XIndexAccess.class, xSheets);
    # xSheet = (XSpreadsheet) UnoRuntime.queryInterface(
    #     XSpreadsheet.class, oIndexSheets.getByIndex(0));

    # # Conversion from Java code :
    # oInterface = (XInterface) oMSF.createInstance(
    #     "com.sun.star.frame.Desktop" );
    # oCLoader = ( XComponentLoader ) UnoRuntime.queryInterface(
    #     XComponentLoader.class, oInterface );
    # PropertyValue [] szEmptyArgs = new PropertyValue [0];
    # aDoc = oCLoader.loadComponentFromURL(
    #     "private:factory/swriter" , "_blank", 0, szEmptyArgs );
    # # -- Python :
    # oCLoader = oMSF.createInstance( "com.sun.star.frame.Desktop" )
    # aDoc = oCLoader.loadComponentFromURL(
    #     "private:factory/swriter", "_blank", 0, () )

    # public static void insertIntoCell(int CellX, int CellY, String theValue,
    #                                   XSpreadsheet TT1, String flag)
    # {
    #     XCell xCell = null;

    #     try {
    #         xCell = TT1.getCellByPosition(CellX, CellY);
    #     } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
    #         System.err.println("Could not get Cell");
    #         ex.printStackTrace(System.err);
    #     }

    #     if (flag.equals("V")) {
    #         xCell.setValue((new Float(theValue)).floatValue());
    #     } else {
    #         xCell.setFormula(theValue);
    #     }

    # }


g_exportedScripts = (HelloWorldWriter, HelloWorldCalc)
