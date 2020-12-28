using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;


namespace drawAxonsSW
{
    /// <summary>
    /// Summary description for drawAxonsSW.
    /// </summary>
    [Guid("bd85d332-72fe-49f4-bc43-3430f4de4af3"), ComVisible(true)]
    [SwAddin(
        Description = "drawAxonsSW - a program designed to draw the axons for a branch of the vestibular nerve",
        Title = "drawAxonsSW",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
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
            int cmdIndex0, cmdIndex1;
            string Title = "Draw Axons", ToolTip = "Functions to Draw Axons";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("drawAxonsSW.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("drawAxonsSW.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("drawAxonsSW.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("drawAxonsSW.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            //cmdIndex0 = cmdGroup.AddCommandItem2("CreateCube", -1, "Create a cube", "Create cube", 0, "CreateCube", "", mainItemID1, menuToolbarOption);
            cmdIndex0 = cmdGroup.AddCommandItem2("Add Reference Points", -1, "Find faces, add reference points to each face w/o repeats, save point and corresponding face data to be used later", "Add Reference Points", 0, "AddRefPts", "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("Add Centerline Slpine", -1, "Add centerline spline using generated refernece points", "Add Centerline Spline", 2, "CreateCube", "", mainItemID2, menuToolbarOption);
            //cmdIndex1 = cmdGroup.AddCommandItem2("Add Centerline Slpine", -1, "Display sample property manager", "Show PMP", 2, "ShowPMP", "EnablePMP", mainItemID2, menuToolbarOption);
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
        public void CreateCube()
        {
            //make sure we have a part open
            string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modDoc = (IModelDoc2)iSwApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                modDoc.InsertSketch2(true);
                modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                //Extrude the sketch
                IFeatureManager featMan = modDoc.FeatureManager;
                featMan.FeatureExtrusion(true,
                    false, false,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    0.1, 0.0,
                    false, false,
                    false, false,
                    0.0, 0.0,
                    false, false,
                    false, false,
                    true,
                    false, false);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }

        #region Add Reference Points
        public void AddRefPts()
        {
            // THIS IS A TEST I MADE A CHANGE
            //make sure we have a part open
            string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modelDoc = (IModelDoc2)iSwApp.IActiveDoc;
                IModelDocExtension swModelDocExt = (IModelDocExtension)modelDoc.Extension;
                IFeatureManager featMgr = modelDoc.FeatureManager;
                ISelectionMgr selMgr = (ISelectionMgr)modelDoc.SelectionManager;
                //ISelectData selData = (ISelectData)selMgr.CreateSelectData();
                Feature feature = (Feature)modelDoc.FirstFeature();
                SketchManager skMgr = (SketchManager)modelDoc.SketchManager;

                (IBodyFolder bodyFolder, List<Feature> bodiesList) = GetBodyFolder(feature);
                Debug.Print(bodyFolder.GetBodyCount().ToString());
                object[] Bodies = (object[])bodyFolder.GetBodies();

                
                
                //Input: bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr
                //Output: refPoints, rpXYZ, rpFeats
                SelectionSet bodySet = GenerateBodySet(bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr);

                (List<MathPoint> refPoints, List<double[]> rpXYZ, List<IFeature> rpFeats, List<double[]> fBoxPts) = CreatePoints(bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr);

                SelectionSet rpSet = GenerateRPSet(swModelDocExt, featMgr, modelDoc, selMgr, refPoints, rpXYZ, rpFeats);
                System.Windows.Forms.MessageBox.Show("A total of " + refPoints.Count + " reference points were created!");

                (SketchSegment splineG, Feature splineFeat) = GenSpline(modelDoc, selMgr, skMgr, rpXYZ);

                List<Feature> refPlanesFeat = GenRefPlanes(modelDoc, selMgr, swModelDocExt, featMgr, rpFeats, splineG, splineFeat);

                List<Feature> coordSystems = GenRefAxis(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlanesFeat);

                Feature combineFeat = GenCombine(modelDoc, swModelDocExt, selMgr, featMgr, bodiesList);

                List<Sketch> boundryCurves = GenBoundarySketches(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlanesFeat, bodyFolder, fBoxPts, skMgr);

                if (selMgr.GetSelectedObjectCount() > 0)
                {
                    selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                    modelDoc.ClearSelection2(true);
                }

                

            }
            else
            {

            }

                

        }

        public List<Sketch> GenBoundarySketches(ISelectionMgr selMgrIN, IModelDoc2 modelDocIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, List<IFeature> rpFeatsIN, List<Feature> refPlanesFeatIN, IBodyFolder bodyFolderIN, List<double[]> fBoxPtsIN, SketchManager skMgrIN)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            object[] bodies = bodyFolderIN.GetBodies();
            Body2 swBody = (Body2)bodies[0];
            object[] feats = (object[])swBody.GetFeatures();
            //Debug.Print("     Number of features in body #" + (int)(0 + 1) + ": " + (int)swBody.GetFeatureCount());
            Feature combineSB = (Feature)feats[0];
            List<Sketch> crossSections = new List<Sketch>();
            //Debug.Print("       Name of feature: " + (string)combineSB.Name);

            object BoxFeatureArrayRP = null;
            double[] BoxFeatureDblArrayRP = new double[7];
            double[] BoxFeatureDblArrayF = new double[7];
            bool status;

            //Debug.Print("Name: " + combineFeat.Name + " Type: " + combineFeat.GetTypeName2());
            //Debug.Print("Name: " + combineSB.Name + " Type: " + combineSB.GetTypeName2());
            status = refPlanesFeatIN[1].GetBox(ref BoxFeatureArrayRP);
            BoxFeatureDblArrayRP = (double[])BoxFeatureArrayRP;
            //Debug.Print("  Pt1 = " + "(" + BoxFeatureDblArrayRP[0] * 1000.0 + ", " + BoxFeatureDblArrayRP[1] * 1000.0 + ", " + BoxFeatureDblArrayRP[2] * 1000.0 + ") mm");
            //Debug.Print("  Pt2 = " + "(" + BoxFeatureDblArrayRP[3] * 1000.0 + ", " + BoxFeatureDblArrayRP[4] * 1000.0 + ", " + BoxFeatureDblArrayRP[5] * 1000.0 + ") mm");

            //Debug.Print("  Pt1 = " + "(" + fBoxPts[0][0] * 1000.0 + ", " + fBoxPts[0][1] * 1000.0 + ", " + fBoxPts[0][2] * 1000.0 + ") mm");
            //Debug.Print("  Pt2 = " + "(" + fBoxPts[0][3] * 1000.0 + ", " + fBoxPts[0][4] * 1000.0 + ", " + fBoxPts[0][5] * 1000.0 + ") mm");
            if (Enumerable.SequenceEqual(BoxFeatureDblArrayRP, fBoxPtsIN[0]))
            {

            }
            else
            {
                
                for (int q = 1; q <= (refPlanesFeatIN.Count - 1); q++)
                {
                    status = swModelDocExtIN.SelectByID2(refPlanesFeatIN[q].Name, "PLANE", 0, 0, 0, false, 0, null, 0);
                    skMgrIN.InsertSketch(true);
                    modelDocIN.ClearSelection2(true);
                    modelDocIN.Sketch3DIntersections();
                    status = swModelDocExtIN.SelectByID2(combineSB.Name, "SOLIDBODY", 0, 0, 0, true, 0, null, 0);
                    modelDocIN.Sketch3DIntersections();
                    crossSections.Add(skMgrIN.ActiveSketch);
                    modelDocIN.ClearSelection2(true);
                    modelDocIN.ClearSelection2(true);
                    skMgrIN.InsertSketch(true);

                }
            }
            return (crossSections);
        }

        public List<Feature> GenRefAxis(ISelectionMgr selMgrIN, IModelDoc2 modelDocIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, List<IFeature> rpFeatsIN, List<Feature> refPlanesFeatIN)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            List<Feature> coordSys = new List<Feature>();
            bool boo;
            for (int s = 0; s <= (rpFeatsIN.Count() - 1); s++)
            {
                boo = swModelDocExtIN.SelectByID2(rpFeatsIN[s].Name, "DATUMPOINT", 0, 0, 0, true, 0, null, 0);
                boo = swModelDocExtIN.SelectByID2(refPlanesFeatIN[s].Name, "PLANE", 0, 0, 0, true, 0, null, 0);
                modelDocIN.ClearSelection2(true);
                boo = swModelDocExtIN.SelectByID2(rpFeatsIN[s].Name, "DATUMPOINT", 0, 0, 0, false, 1, null, 0);
                boo = swModelDocExtIN.SelectByID2(refPlanesFeatIN[s].Name, "PLANE", 0, 0, 0, true, 8, null, 0);
                Feature coordS = featMgrIN.InsertCoordinateSystem(false, false, false);
                coordSys.Add(coordS);
            }
            return (coordSys);
        }
        public Feature GenCombine(IModelDoc2 modelDocIN, IModelDocExtension swModelDocExtIN, ISelectionMgr selMgrIN, IFeatureManager featMgrIN, List<Feature> bodiesListIN)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            bool boolSt;
            foreach (Feature bf in bodiesListIN)
            {
                //Debug.Print("Name: " + bf.Name + ", Type: " + bf.GetTypeName());
                boolSt = swModelDocExtIN.SelectByID2(bf.Name, "SOLIDBODY", 0, 0, 0, true, 2, null, 0);
            }
            Feature combineFeat = featMgrIN.InsertCombineFeature(15903, null, null);
            return (combineFeat);
        }

        public List<Feature> GenRefPlanes(IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, List<IFeature> rpFeatsIN, SketchSegment splineGIN, Feature splineFeatIN)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            bool bools;
            List<Feature> refPlanes = new List<Feature>();
            foreach (IFeature fs in rpFeatsIN)
            {
                bools = swModelDocExtIN.SelectByID2(splineGIN.GetName() + "@" + splineFeatIN.Name, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                bools = swModelDocExtIN.SelectByID2(fs.Name, "DATUMPOINT", 0, 0, 0, true, 1, null, 0);
                Debug.Print("Num of selected" + selMgrIN.GetSelectedObjectCount());
                object myRefPlane = featMgrIN.InsertRefPlane(2, 0, 4, 0, 0, 0);
                Feature refPlane = (Feature)myRefPlane;
                refPlanes.Add(refPlane);
                modelDocIN.ClearSelection2(true);
            }
            return (refPlanes);
        }

        public (SketchSegment, Feature) GenSpline(IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, SketchManager skMgrIN, List<double[]> rpXYZIN)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            double[] points = new double[rpXYZIN.Count * 3];
            object pointArray = null;
            for (int i = 0; i <= rpXYZIN.Count - 1; i++)
            {
                points[i + 2 * i] = rpXYZIN[i][0];
                points[i + 1 + 2 * i] = rpXYZIN[i][1];
                points[i + 2 + 2 * i] = rpXYZIN[i][2];

            }
            pointArray = points;
            skMgrIN.Insert3DSketch(true);
            Sketch sk = skMgrIN.ActiveSketch;
            Feature splineSketchFeat = (Feature)sk;
            SketchSegment swSketchSegment = (SketchSegment)skMgrIN.CreateSpline2((pointArray), true);
            modelDocIN.ClearSelection2(true);
            skMgrIN.Insert3DSketch(true);
            return (swSketchSegment,splineSketchFeat);
        }

        public SelectionSet GenerateRPSet(IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, List<MathPoint> refPointsIN, List<double[]> rpXYZIN, List<IFeature> rpFeatsIN)
        {
            int markI2 = 101;
            int errors2 = 0;
            bool boolstatus2;
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            for (int i = 0; i <= (rpFeatsIN.Count - 1); i++)
            {
                boolstatus2 = rpFeatsIN[i].Select2(true, 0);
                //Debug.Print("Num selected = " + selMgrIN.GetSelectedObjectCount());
                //Debug.Print("Mark = " + selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()));
                if (selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()) == 0)
                {
                    selMgrIN.SetSelectedObjectMark(selMgrIN.GetSelectedObjectCount(), markI2++, 0);
                    //Debug.Print("Mark = " + selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()));
                }
            }
            SelectionSet rpSetOUT = (SelectionSet)swModelDocExtIN.SaveSelection(out errors2);
            Debug.Print("  First selection set created (1 = succeeded; 0 = failed)? " + errors2);
            modelDocIN.ClearSelection2(true);
            return (rpSetOUT);

        }

        public (List<MathPoint> refPointsOUT, List<double[]> rpXYZOUT, List<IFeature> rpFeatsOUT, List<double[]> faceBox) CreatePoints(IBodyFolder bodyFolderIN, object[] BodiesIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, IModelDoc2 modelDocIN, ISelectionMgr selMgrIN)
        {
            bool boolstatus;
            bool addFlag = true;
            List<double[]> faceD = new List<double[]>(); 
            List<double[]> faceBox = new List<double[]>();
            for (int i = 0; i <= (bodyFolderIN.GetBodyCount() - 1); i++)
            {
                IBody2 Body = (IBody2)BodiesIN[i];
                object[] bodyFaces = (object[])Body.GetFaces();
                for (int j = 0; j <= (bodyFaces.Length - 1); j++)
                {
                    IFace2 faceB = (IFace2)bodyFaces[j];
                    //Debug.Print("Number of Edges = " + faceB.GetEdgeCount().ToString() + ", Area of Face = " + faceB.GetArea().ToString());
                    if (faceB.GetEdgeCount() > 20)
                    {
                        double[] ptD = new double[3];
                        double[] pts2 = (double[])faceB.GetBox();

                        //Debug.Print("X1= " + pts2[0] + " Y1= " + pts2[1] + " Z1= " + pts2[2] + " X2= " + pts2[3] + " Y2= " + pts2[4] + " Z2= " + pts2[5]);
                        double xD = (pts2[3] - pts2[0]) / 2;
                        double yD = (pts2[4] - pts2[1]) / 2;
                        double zD = (pts2[5] - pts2[2]) / 2;
                        swModelDocExtIN.SelectByID2("", "FACE", pts2[0] + xD, pts2[1] + yD, pts2[2] + zD, false, 0, null, 0);
                        ptD[0] = pts2[0] + xD;
                        ptD[1] = pts2[1] + yD;
                        ptD[2] = pts2[2] + zD;
                        if (faceD.Count == 0)
                        {
                            faceD.Add(ptD);
                            faceBox.Add(pts2);
                            object[] centerLinePts = (object[])featMgrIN.InsertReferencePoint(4, 0, 0.01, 1);
                        }
                        else
                        {
                            addFlag = true;
                            for (int k = 0; k <= (faceD.Count - 1); k++)
                            {
                                if (System.Linq.Enumerable.SequenceEqual(faceD[k], ptD))
                                {
                                    addFlag = false;
                                    k = (faceD.Count - 1);
                                }
                                else
                                {
                                    addFlag = true;
                                }
                            }
                            if (addFlag)
                            {
                                faceD.Add(ptD);
                                faceBox.Add(pts2);
                                object[] centerLinePts = (object[])featMgrIN.InsertReferencePoint(4, 0, 0.01, 1);
                            }
                        }
                        
                    }
                }
            }

            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }

            IFeature feature2 = (IFeature)modelDocIN.FirstFeature();
            bool firstPass = true;
            List<MathPoint> refPointsOUT = new List<MathPoint>();
            List<double[]> rpXYZOUT = new List<double[]>();
            List<IFeature> rpFeatsOUT = new List<IFeature>();
            while (feature2 != null)
            {
                string FeatType = feature2.Name;
                string FeatTypeName = feature2.GetTypeName2();
                //Debug.Print(" " + FeatType + " [" + FeatTypeName + "]");
                if (FeatTypeName == "RefPoint")
                {

                    IRefPoint rp = (IRefPoint)feature2.GetSpecificFeature2();
                    MathPoint mp = rp.GetRefPoint();
                    double[] mp2 = mp.ArrayData;
                    if (rpXYZOUT.Count == 0)
                    {
                        refPointsOUT.Add(mp);
                        rpXYZOUT.Add(mp2);
                        rpFeatsOUT.Add(feature2);
                    }
                    else
                    {
                        addFlag = true;
                        for (int i = 0; i <= (rpXYZOUT.Count - 1); i++)
                        {
                            if (Enumerable.SequenceEqual(rpXYZOUT[i], mp2))
                            {
                                boolstatus = swModelDocExtIN.SelectByID2(FeatType, "DATUMPOINT", 0, 0, 0, true, 0, null, 0);
                                addFlag = false;
                                i = (rpXYZOUT.Count - 1);
                            }
                            else
                            {
                                addFlag = true;
                            }
                        }
                        if (addFlag)
                        {
                            refPointsOUT.Add(mp);
                            rpXYZOUT.Add(mp2);
                            rpFeatsOUT.Add(feature2);
                        }
                    }
                }
                feature2 = (IFeature)feature2.GetNextFeature();
                if ((feature2 == null) && (firstPass == true))
                {
                    modelDocIN.EditDelete();
                    feature2 = (IFeature)modelDocIN.FirstFeature();
                    firstPass = false;
                    refPointsOUT.Clear();
                    rpXYZOUT.Clear();
                    rpFeatsOUT.Clear();
                }
            }
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                modelDocIN.EditDelete();
            }

            return (refPointsOUT, rpXYZOUT, rpFeatsOUT, faceBox);
        }

        public SelectionSet GenerateBodySet(IBodyFolder bodyFolderIN, object[] BodiesIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, IModelDoc2 modelDocIN, ISelectionMgr selMgrIN)
        {
            int markI = 1;
            int errors = 0;
            bool boolstatus;
            SelectData swSelData = (SelectData)selMgrIN.CreateSelectData();
            for (int i = 0; i <= (bodyFolderIN.GetBodyCount() - 1); i++)
            {
                Body2 Body = (Body2)BodiesIN[i];
                boolstatus = Body.Select2(true, swSelData);
                //Debug.Print("Num selected = " + selMgrIN.GetSelectedObjectCount());
                  //  Debug.Print("Mark = " + selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()));
                    if (selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()) == 0)
                    {
                        selMgrIN.SetSelectedObjectMark(selMgrIN.GetSelectedObjectCount(), markI++, 0);
                        //Debug.Print("Mark = " + selMgrIN.GetSelectedObjectMark(selMgrIN.GetSelectedObjectCount()));
                    }
            }
            SelectionSet bodySetOUT = (SelectionSet)swModelDocExtIN.SaveSelection(out errors);
            Debug.Print("  First selection set created (1 = succeeded; 0 = failed)? " + errors);
            modelDocIN.ClearSelection2(true);
            return (bodySetOUT);
        }

        public (IBodyFolder, List<Feature>) GetBodyFolder(Feature swFeature)
        {
            bool go = true;
            List<Feature> bodies = new List<Feature>();
            IBodyFolder swBodyFolder = default(IBodyFolder);
            while (go)
            {
                string FeatType = swFeature.Name;
                string FeatTypeName = swFeature.GetTypeName2();
                Debug.Print(" " + FeatType + " [" + FeatTypeName + "]");
                if (FeatTypeName == "SolidBodyFolder")
                {
                    Debug.Print(" Returing BodyFolder variable ");
                    swBodyFolder = (IBodyFolder)swFeature.GetSpecificFeature2();                    
                }
                if (FeatTypeName == "BaseBody")
                {
                    bodies.Add(swFeature);
                }
                swFeature = (Feature)swFeature.GetNextFeature();
                if (swFeature == null)
                {
                    go = false;
                    return (swBodyFolder, bodies);
                    
                }
            }
            return (null, null);
        }
        #endregion












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
