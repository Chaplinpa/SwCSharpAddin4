using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualBasic;
using Microsoft.CSharp;
using System.Windows.Forms;


namespace SwCSharpAddin4
{
    /// <summary>
    /// Summary description for SwCSharpAddin4.
    /// </summary>
    [Guid("c41f93d4-e532-4426-b035-3aa6960bdf71"), ComVisible(true)]
    [SwAddin(
        Description = "General purpose SOLIDWORKS Add-In.",
        Title = "SwCSharpAddin4",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        internal static ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int mainItemID4 = 3;
        public const int flyoutGroupID = 91;

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0, cmdIndex1, cmdIndex2, cmdIndex3;
            string Title = "C# Addin", ToolTip = "C# Addin";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[4] { mainItemID1, mainItemID2, mainItemID3, mainItemID4 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("SwCSharpAddin4.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("SwCSharpAddin4.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("SwCSharpAddin4.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("SwCSharpAddin4.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            // cmdIndex0 = cmdGroup.AddCommandItem2("CreateCube", -1, "Create a cube", "Create cube", 0, "CreateCube", "", mainItemID1, menuToolbarOption);
            cmdIndex0 = cmdGroup.AddCommandItem2("SaveDrawingV2", -1, "Saved the Drawing.", "Save DrawingV2", 0, "SaveDrawingV2", "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("OpenInExplorer", -1, "Browse to the selected file in Windows Explorer", "OpenInExplorer", 1, "OpenInExplorer", "", mainItemID2, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddCommandItem2("OpenDrawing", -1, "Opens Drawing",  "OpenDrawing", 2, "OpenDrawing", "", mainItemID3, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2("DrawingConfigSwitch", -1, "Purges Suppressed Components", "DrawingConfigSwitch", 3, "DrawingConfigSwitch", "", mainItemID4, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;



            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");


            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[3];
                    int[] TextType = new int[3];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);

                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[2] = cmdGroup.ToolbarId;

                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);



                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);

                }

            }
            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }

        #endregion

        #region UI Callbacks
    
        public void CreateDXF()
        {
            //This callback is a utility to create a DXF flat pattern file for a sheet metal part.


        }
        public void SaveDrawingV2()
        {

            //This callback is for a utility to save an un-saved drawing of a part with a BPA part number into the correct Vault directory with the correct file name.
            
            DrawingDoc swDraw = iSwApp.ActiveDoc;
            SolidWorks.Interop.sldworks.View swView = swDraw.IGetFirstView();
            swView = swView.GetNextView();
            ModelDoc2 swDrawModel = swView.ReferencedDocument;
            int errors = 0;
            int warnings = 0;
            String configName = swView.ReferencedConfiguration;
            String drawingNumber = swDrawModel.GetCustomInfoValue(configName, "Drawing Number");
            //MessageBox.Show("The drawing number is: " + drawingNumber);
            

            if (drawingNumber == "" || drawingNumber == " ")
            {
                MessageBox.Show("Please fill out the drawing number field on the model's file data card before proceeding.","Missing Drawing Number" , MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            String modelPath = swDrawModel.GetPathName().ToString();
           MessageBox.Show("The model's path is: " + modelPath);
            if (modelPath == "" || modelPath == " ")
            {
                MessageBox.Show("Please save the model before proceeding", "Save the Model", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            swDrawModel = iSwApp.ActiveDoc;
            String swDrawPath = swDrawModel.GetPathName().ToString();
            //MessageBox.Show("The drawing path is: " + swDrawPath);
            if (swDrawPath != "")
            {
                MessageBox.Show("The drawing has already been saved as '" + swDrawPath + "' in the Vault.", "Drawng Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                
                string[] pathDir = modelPath.Split('\\');
                int pathDirLength = pathDir.Length;
                string path = pathDir[0] + "\\" + pathDir[1] + "\\" + pathDir[2] + "\\" + pathDir[3] + "\\" + pathDir[4] +"\\";
                string dwgPath = path + drawingNumber + ".SLDDRW";
                string dialogText = "The drawing path is: \n" + dwgPath + "\nContinue?";
                DialogResult result = MessageBox.Show(dialogText, "Save Drawing?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result == DialogResult.OK)
                {
                    ModelDoc2 Model = SwApp.ActiveDoc;
                    Boolean Save;
                    Save = Model.SaveAs4(dwgPath, 0, 1, errors, warnings);
                }

                }

            
            //ModelDocExtension swModExt = default(ModelDocExtension);
            //Return the filepath of the referenced model in the parent view in the active drawing document.
            //string swModel = swView.GetReferencedModelName();


        }
        public void OpenInExplorer()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc;
            string swModelPath = swModel.GetPathName();
            swModelPath = "/select, " + swModelPath;

            System.Diagnostics.Process Process = new System.Diagnostics.Process();
            Process.StartInfo.UseShellExecute = true;
            Process.StartInfo.FileName = @"explorer";
            Process.StartInfo.Arguments = swModelPath;
            Process.Start();

        }
        public void PurgeUnused()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc;
            Configuration swConf = swModel.GetActiveConfiguration();
            var cfgNames = swModel.GetConfigurationNames();
            ConfigurationManager swCfgMgr = swModel.ConfigurationManager;
           
            Component2 swRootComp = swConf.GetRootComponent3(true);
            Boolean bRet = true;
            int i;
            string configName = "";
            object ConfigValues = new object();
            object ConfigParams = new object();
            var vChildComp = swRootComp.GetChildren();

            for (i = 0; i <= vChildComp.GetUpperBound(0); i++)
            {
                Component2 swChildComp = vChildComp[i];
                
                bRet = swCfgMgr.GetConfigurationParams(configName, out ConfigParams, out ConfigValues);    

                //if (swChildComp.GetSuppression() == 0)
                //{
                //    MessageBox.Show(swChildComp.Name2 + " is suppressed");
                //    swChildComp.Select(true);
                //    swModel.EditDelete();

                //}
            }
                       
            



        }
        public void OpenDrawing()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc;
            //SelectionMgr selMgr = new SelectionMgr();
            int iErrors = 0;
            int iWarnings = 0;
            Component2 swComponent;
            String swComponentPath;
            
            SelectionMgr selMgr = default(SelectionMgr);
            selMgr = (SelectionMgr)swModel.SelectionManager;
            swComponent = selMgr.GetSelectedObjectsComponent4(1, -1);

            try
            {
                swComponentPath = swComponent.GetPathName();
            }
            catch (NullReferenceException)
            {
                swComponentPath = swModel.GetPathName();
                //swComponentPath = swComponent.GetPathName();
            }
            var swDrawingArr = swComponentPath.Split('.');
            String swDrawPath = swDrawingArr[0] + ".SLDDRW";
            iSwApp.OpenDoc6(swDrawPath, 3, 1,"", iErrors, iWarnings);

            //MessageBox.Show(swComponentPath.ToString());
            



        }
        public void DrawingConfigSwitch()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc;
            DrawingDoc swDraw;
            ModelDoc2 swRefPart;
            string refPartName = "";
            string refPartCfg = "";
            SolidWorks.Interop.sldworks.View swView;
            

            int Type = swModel.GetType();
            if (Type != 3)
            {
                MessageBox.Show("Please open a drawing document.");
                goto LineEnd;
            }

            swDraw = (DrawingDoc)swModel;
            swView = swDraw.GetFirstView() as SolidWorks.Interop.sldworks.View;

            while (swView != null)
            {
                
                refPartName = swView.GetReferencedModelName();
                refPartCfg = swView.ReferencedConfiguration;
                if (refPartName != "")
                {
                    goto Line1;
                }
                swView = swView.GetNextView();
            }
        Line1:;


            swRefPart = swView.ReferencedDocument;
            string[] cfgNames = swRefPart.GetConfigurationNames();
            frmDrawConfigSwitch configForm = new frmDrawConfigSwitch(cfgNames, refPartCfg);
            configForm.Show();

            LineEnd:;
        

        }

        internal static void drwChangeConfig(string cfgName)
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc;
            DrawingDoc swDraw;
            bool retVal;
            SolidWorks.Interop.sldworks.View swView;

            swDraw = (DrawingDoc)swModel;
            swView = swDraw.GetFirstView() as SolidWorks.Interop.sldworks.View;

            while (swView != null)
            {
                retVal = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swView.ReferencedConfiguration = cfgName;
                swView = swView.GetNextView();
            }
            swDraw.ForceRebuild();     
            
        }


        public void ShowPMP()
    {
        if (ppage != null)
            ppage.Show();
    }

    public int EnablePMP()
    {
        if (iSwApp.ActiveDoc != null)
            return 1;
        else
            return 0;
    }

    public void FlyoutCallback()
    {
        FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
        flyGroup.RemoveAllCommandItems();

        flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

    }
    public int FlyoutEnable()
    {
        return 1;
    }

    public void FlyoutCommandItem1()
    {
        iSwApp.SendMsgToUser("Flyout command 1");
    }

    public int FlyoutEnableCommandItem1()
    {
        return 1;
    }
    #endregion

    #region Event Methods
    public bool AttachEventHandlers()
    {
        AttachSwEvents();
        //Listen for events on all currently open docs
        AttachEventsToAllDocuments();
        return true;
    }

    private bool AttachSwEvents()
    {
        try
        {
            SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
            SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
            SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
            SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
            SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
            return false;
        }
    }



    private bool DetachSwEvents()
    {
        try
        {
            SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
            SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
            SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
            SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
            SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
            return false;
        }

    }

    public void AttachEventsToAllDocuments()
    {
        ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
        while (modDoc != null)
        {
            if (!openDocs.Contains(modDoc))
            {
                AttachModelDocEventHandler(modDoc);
            }
            modDoc = (ModelDoc2)modDoc.GetNext();
        }
    }

    public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
    {
        if (modDoc == null)
            return false;

        DocumentEventHandler docHandler = null;

        if (!openDocs.Contains(modDoc))
        {
            switch (modDoc.GetType())
            {
                case (int)swDocumentTypes_e.swDocPART:
                    {
                        docHandler = new PartEventHandler(modDoc, this);
                        break;
                    }
                case (int)swDocumentTypes_e.swDocASSEMBLY:
                    {
                        docHandler = new AssemblyEventHandler(modDoc, this);
                        break;
                    }
                case (int)swDocumentTypes_e.swDocDRAWING:
                    {
                        docHandler = new DrawingEventHandler(modDoc, this);
                        break;
                    }
                default:
                    {
                        return false; //Unsupported document type
                    }
            }
            docHandler.AttachEventHandlers();
            openDocs.Add(modDoc, docHandler);
        }
        return true;
    }

    public bool DetachModelEventHandler(ModelDoc2 modDoc)
    {
        DocumentEventHandler docHandler;
        docHandler = (DocumentEventHandler)openDocs[modDoc];
        openDocs.Remove(modDoc);
        modDoc = null;
        docHandler = null;
        return true;
    }

    public bool DetachEventHandlers()
    {
        DetachSwEvents();

        //Close events on all currently open docs
        DocumentEventHandler docHandler;
        int numKeys = openDocs.Count;
        object[] keys = new Object[numKeys];

        //Remove all document event handlers
        openDocs.Keys.CopyTo(keys, 0);
        foreach (ModelDoc2 key in keys)
        {
            docHandler = (DocumentEventHandler)openDocs[key];
            docHandler.DetachEventHandlers(); //This also removes the pair from the hash
            docHandler = null;
        }
        return true;
    }
    #endregion

    #region Event Handlers
    //Events
    public int OnDocChange()
    {
        return 0;
    }

    public int OnDocLoad(string docTitle, string docPath)
    {
        return 0;
    }

    int FileOpenPostNotify(string FileName)
    {
        AttachEventsToAllDocuments();
        return 0;
    }

    public int OnFileNew(object newDoc, int docType, string templateName)
    {
        AttachEventsToAllDocuments();
        return 0;
    }

    public int OnModelChange()
    {
        return 0;
    }

        

        #endregion

    }


      

}


    