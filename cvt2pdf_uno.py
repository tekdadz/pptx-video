import uno
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue


def open_pptx(file_path):
    # Establish a connection to the LibreOffice process
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except NoConnectException:
        print("Failed to connect to LibreOffice.")
        return

    # Access the central desktop object
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

    # Define properties for loading the document
    # Here, you could specify properties for the document opening mode
    # load_props = tuple(PropertyValue(Name="Hidden", Value=True) for i in range(1))

    # Load the document
    document = desktop.loadComponentFromURL(uno.systemPathToFileUrl(file_path), "_blank", 0, ())
    if document:
        print("Presentation opened successfully.")

        out_props = (
            PropertyValue(Name="FilterName", Value="impress_pdf_Export"),
            PropertyValue(Name="Overwrite", Value=True)
        )
        document.storeToURL(uno.systemPathToFileUrl("E:\\python\\pptx-video\\test.pdf"), out_props)

        # When done, close the document
        document.close(True)
    else:
        print("Failed to open the document.")


# Example usage
open_pptx("E:\\python\\pptx-video\\test.pptx")

