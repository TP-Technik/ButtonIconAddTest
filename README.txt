Hi,
For refferences I used the COM references to stdole (OLE Automation) and Autodesk Inventor Object libraty
Alternatively you can directly reference them from the local Autodesk Inventor file direction (Autodesk.Inventor.Interop.dll and stdole.dll) both found within the folder structure: "Autodesk\Inventor XXXX\Bin..."

If you want to use Resource you will need to add references to System.Windows.Forms to be able to handle Bitmap and PNG files.
