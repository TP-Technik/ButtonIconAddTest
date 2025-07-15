using Inventor;
using IPictureDisp = stdole.IPictureDisp;
using Bitmap = System.Drawing.Bitmap;

public class Programm
{

    private static String AddInClientID = "myAddInID";

    private static RibbonTab? myCustomTab;
    private static RibbonPanel? createPannel;
    private static ButtonDefinition? myCustomButton;
    private static CommandControl? myCustomButtonControl;

    [STAThread]
    public static void Main()
    {
        var invApp = Marshal2.GetActiveObject("Inventor.Application") as Application;

        var ribbonRefference = invApp.UserInterfaceManager.Ribbons["Assembly"];

        // Create custom Tab:
        myCustomTab = ribbonRefference.RibbonTabs.Add(
            "myCustomTab",
            "InternalName1",
            AddInClientID);

        // Create custom Pannel:
        createPannel = myCustomTab.RibbonPanels.Add(
            "createPannel",
            "InternalName2",
            AddInClientID);

        var controlDefs = invApp.CommandManager.ControlDefinitions;


        Bitmap standardBmp = new Bitmap(@"C:\...\Plug16.bmp");
        Bitmap largeBmp = new Bitmap(@"C:\...\Plug32.bmp");

        IPictureDisp standardPict = PictureDispConverter.ToIPictureDisp(standardBmp);
        IPictureDisp largePict = PictureDispConverter.ToIPictureDisp(largeBmp);


        // Create custom Button:
        myCustomButton = controlDefs.AddButtonDefinition(
            "myCustomButton",
            "InternalName3",
            CommandTypesEnum.kEditMaskCmdType,
            StandardIcon: standardPict,
            LargeIcon: largePict);

        // If you break hear you can Inspect the error of the myCustomButton Icons, by hovering over it:
        // myCustomButton.StandardIcon = standardPict;
        // myCustomButton.LargeIcon = largePict;


        myCustomButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(myFunction);

        // Add Button:
        myCustomButtonControl = createPannel.CommandControls.AddButton(myCustomButton);

    }

    private static void myFunction(NameValueMap context)
    {
        // Do something
    }

}


