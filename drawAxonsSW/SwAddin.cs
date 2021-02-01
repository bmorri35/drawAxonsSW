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
using Extreme.Mathematics;
using Extreme.Mathematics.Random;
using Extreme.Statistics.Distributions;
using Extreme.Mathematics.LinearAlgebra;
using System.Data;
using ExtensionMethods;
using System.Windows.Forms;

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
                int its = 0;

                IModelDoc2 modDoc = (IModelDoc2)iSwApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);
                IModelDocExtension swModelDocExt = (IModelDocExtension)modDoc.Extension;
                IFeatureManager featMgr = modDoc.FeatureManager;
                ISelectionMgr selMgr = (ISelectionMgr)modDoc.SelectionManager;
                Feature feature = (Feature)modDoc.FirstFeature();
                SketchManager skMgr = (SketchManager)modDoc.SketchManager;
                List<Feature> RefPlanes = new List<Feature>();
                List<Feature> sketchFeatures = new List<Feature>();
                List<Matrix<double>> RotationMatricies = new List<Matrix<double>>();
                List<Matrix<double>> TranslationMatricies = new List<Matrix<double>>();
                List<Sketch> sketchs = new List<Sketch>();
                List<MathPoint> RefMathPt = new List<MathPoint>();
                List<Matrix<double>> RefPtsMatGlobal = new List<Matrix<double>>();
                List<Matrix<double>> RefPtsMatSketch = new List<Matrix<double>>();
                List<List<double[]>> Axons = new List<List<double[]>>();
                List<List<double[]>> AllvertXYZ = new List<List<double[]>>();
                List<List<Matrix<double>>> AllvertexGlobal = new List<List<Matrix<double>>>();
                List<List<Matrix<double>>> AllvertexInSkPlane = new List<List<Matrix<double>>>();
                List<List<double[]>> AllScaledVertsXList = new List<List<double[]>>();
                List<List<double[]>> AllScaledVertsYList = new List<List<double[]>>();
                List<List<double[]>> AllScaledVertsZList = new List<List<double[]>>();
                List<List<double>> AllvertAngList = new List<List<double>>();

                UserUnit unit = (UserUnit)modDoc.GetUserUnit(0);
                double cf = unit.GetConversionFactor();

                NormalDistribution norm = new NormalDistribution();
                ContinuousUniformDistribution uni = new ContinuousUniformDistribution(0, 360);
                MersenneTwister random = new MersenneTwister();

                MLApp.MLApp matlab = new MLApp.MLApp();

                bool bools;
                Body2 body = default(Body2);
                object myRefPlane = new object();
                Feature refPlane;
                Sketch sk;
                Feature fe;
                
                while (its < 2)
                {
                    Matrix<double> rotationMatrix = null;
                    Matrix<double> translationMatrix = null;
                    Matrix<double> refPXYZGlobal = null;
                    Matrix<double> refPXYZInSkPlane = null;
                    List<Vertex> verts = new List<Vertex>();
                    List<double[]> vertXYZ = new List<double[]>();
                    List<Matrix<double>> vertexGlobal = new List<Matrix<double>>();
                    List<Matrix<double>> vertexInSkPlane = new List<Matrix<double>>();
                    List<double[]> ScaledVertsXList = new List<double[]>();
                    List<double[]> ScaledVertsYList = new List<double[]>();
                    List<double[]> ScaledVertsZList = new List<double[]>();
                    List<double> vertAngList = new List<double>();
                    if (its == 0)
                    {
                        modDoc.InsertSketch2(true);
                        modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                        #region Make Box
                        featMgr.FeatureExtrusion(true,
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
                        #endregion

                        Feature feat = (Feature)modDoc.FirstFeature();
                        object[] faceObjs = null;
                        Face2 swFace = default(Face2);
                        List<Face2> swFaces = new List<Face2>();
                        List<double[]> boxes = new List<double[]>();
                        #region Save Body, Get Faces and Bounding Boxes
                        while (feat != null)
                        {
                            Debug.Print("Name: " + feat.Name + "Type: " + feat.GetTypeName2());
                            if (feat.Name == "Solid Bodies")
                            {
                                BodyFolder swBodyFolder = (BodyFolder)feat.GetSpecificFeature2();
                                object[] bodies = (object[])swBodyFolder.GetBodies();
                                body = (Body2)bodies[0];
                                faceObjs = (object[])body.GetFaces();
                            }
                            feat = feat.GetNextFeature();
                        }

                        for (int i = faceObjs.GetLowerBound(0); i <= faceObjs.GetUpperBound(0); i++)
                        {
                            swFace = (Face2)faceObjs[i];
                            double[] ptD = new double[3];
                            double[] pts2 = (double[])swFace.GetBox();

                            //Debug.Print("X1= " + pts2[0] + " Y1= " + pts2[1] + " Z1= " + pts2[2] + " X2= " + pts2[3] + " Y2= " + pts2[4] + " Z2= " + pts2[5]);
                            double xD = (pts2[3] - pts2[0]) / 2;
                            double yD = (pts2[4] - pts2[1]) / 2;
                            double zD = (pts2[5] - pts2[2]) / 2;
                            ptD[0] = pts2[0] + xD;
                            ptD[1] = pts2[1] + yD;
                            ptD[2] = pts2[2] + zD;
                            //swModelDocExt.SelectByID2("", "FACE", pts2[0] + xD, pts2[1] + yD, pts2[2] + zD, false, 0, null, 0);
                            boxes.Add(ptD);
                            swFaces.Add(swFace);

                            //Debug.Print("Area of face " + i + " of midsurface feature: " + swFace.GetArea());
                        }
                        #endregion
                        #region Create and Save Ref Point
                        bools = swModelDocExt.SelectByID2("", "FACE", boxes[0][0], boxes[0][1], boxes[0][2], true, 0, null, 0);
                        object[] centerLinePts = (object[])featMgr.InsertReferencePoint(4, 0, 0.01, 1);
                        modDoc.ClearSelection2(true);
                        object clp = centerLinePts[0];
                        Feature comF = (Feature)clp;
                        RefPoint comRefPt = (RefPoint)comF.GetSpecificFeature2();
                        MathPoint comMathPt = comRefPt.GetRefPoint();
                        RefMathPt.Add(comMathPt);
                        #endregion
                        #region Create and Save Ref Plane
                        Debug.Print("Num of selected" + selMgr.GetSelectedObjectCount());
                        bools = swModelDocExt.SelectByID2("", "FACE", boxes[0][0], boxes[0][1], boxes[0][2], true, 0, null, 0);
                        myRefPlane = featMgr.InsertRefPlane(4, 0, 0, 0, 0, 0);
                        refPlane = (Feature)myRefPlane;
                        RefPlanes.Add(refPlane);
                        modDoc.ClearSelection2(true);
                        #endregion
                        #region Create Intersection Sketch and save sketch
                        bools = swModelDocExt.SelectByID2(refPlane.Name, "PLANE", 0, 0, 0, false, 0, null, 0);
                        skMgr.InsertSketch(false);
                        modDoc.ClearSelection2(true);
                        bools = swModelDocExt.SelectByID2(body.Name, "SOLIDBODY", 0, 0, 0, true, 0, null, 0);
                        modDoc.Sketch3DIntersections();
                        sk = skMgr.ActiveSketch;
                        modDoc.ClearSelection2(true);
                        modDoc.ClearSelection2(true);
                        skMgr.InsertSketch(true);
                        sketchs.Add(sk);
                        System.Windows.Forms.MessageBox.Show("A total of ");
                        bools = swModelDocExt.Rebuild(2);
                        fe = swModelDocExt.GetLastFeatureAdded();
                        sketchFeatures.Add(fe);
                        #endregion
                        if (selMgr.GetSelectedObjectCount() > 0)
                        {
                            selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            modDoc.ClearSelection2(true);
                        }
                        #region Get and save Intersection Curve Vert & Angles
                        (rotationMatrix, translationMatrix, refPXYZGlobal, refPXYZInSkPlane, verts, vertXYZ, vertexGlobal, vertexInSkPlane, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, vertAngList) = GetIntersectionCurveVertAngles(sk, comMathPt, cf, selMgr, modDoc, fe, swModelDocExt);
                        RotationMatricies.Add(rotationMatrix);
                        TranslationMatricies.Add(translationMatrix);
                        RefPtsMatGlobal.Add(refPXYZGlobal);
                        RefPtsMatSketch.Add(refPXYZInSkPlane);
                        AllvertXYZ.Add(vertXYZ);
                        AllvertexGlobal.Add(vertexGlobal);
                        AllvertexInSkPlane.Add(vertexInSkPlane);
                        AllScaledVertsXList.Add(ScaledVertsXList);
                        AllScaledVertsYList.Add(ScaledVertsYList);
                        AllScaledVertsZList.Add(ScaledVertsZList);
                        AllvertAngList.Add(vertAngList);

                        double[] xvs = new double[ScaledVertsXList.Count];
                        double[] yvs = new double[ScaledVertsXList.Count];
                        //matlab.Feval("polyxpoly", 2, out inter, refPtXs, refPtYs, newPtXs, newPtYs);
                        //interPts = (object[])inter;
                        for (int qt = 0; qt <= (ScaledVertsXList.Count-1); qt++)
                        {
                            xvs[qt] = ScaledVertsXList[qt][0];
                            yvs[qt] = ScaledVertsYList[qt][0];
                        }
                        
                        System.Windows.Forms.MessageBox.Show("Centroid, X: "+xvs.Average()+" Y: "+yvs.Average());
                        #endregion
                        #region Generate and save seed points
                        bools = swModelDocExt.SelectByID2(fe.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modDoc.EditSketch();

                        int sketchPoints = 1;
                        while (sketchPoints <= 300)
                        {
                            double[] intersectionPointsXY = new double[2];
                            double[] distance = new double[10];
                            double angle;
                            double distance2Test;
                            bool good2Go;
                            int RingWIn;

                            // Generate point for outer third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistWithBounds(distance[3], distance[0], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, true, false, false, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX / cf, dY / cf, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                double[] node = new double[13];
                                List<double[]> axon = new List<double[]>();
                                node[0] = vPoint[0];
                                node[1] = vPoint[1];
                                node[2] = vPoint[2];
                                node[3] = vPointInSketch[0];
                                node[4] = vPointInSketch[1];
                                node[5] = vPointInSketch[2];
                                node[6] = distance2Test / distance[0];
                                node[7] = distance2Test;
                                node[8] = angle;
                                node[9] = RingWIn;
                                node[10] = 10;
                                node[11] = 10;
                                node[12] = 300;
                                axon.Add(node);
                                Axons.Add(axon);


                            }

                            // Generate point for middle third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistWithBounds(distance[6], distance[3], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, false, true, false, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX / cf, dY / cf, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                double[] node = new double[13];
                                List<double[]> axon = new List<double[]>();
                                node[0] = vPoint[0];
                                node[1] = vPoint[1];
                                node[2] = vPoint[2];
                                node[3] = vPointInSketch[0];
                                node[4] = vPointInSketch[1];
                                node[5] = vPointInSketch[2];
                                node[6] = distance2Test / distance[0];
                                node[7] = distance2Test;
                                node[8] = angle;
                                node[9] = RingWIn;
                                node[10] = 10;
                                node[11] = 10;
                                node[12] = 300;
                                axon.Add(node);
                                Axons.Add(axon);
                            }

                            // Generate point for inner third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistCenter(distance[6], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, false, false, true, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX / cf, dY / cf, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                double[] node = new double[13];
                                List<double[]> axon = new List<double[]>();
                                node[0] = vPoint[0];
                                node[1] = vPoint[1];
                                node[2] = vPoint[2];
                                node[3] = vPointInSketch[0];
                                node[4] = vPointInSketch[1];
                                node[5] = vPointInSketch[2];
                                node[6] = distance2Test / distance[0];
                                node[7] = distance2Test;
                                node[8] = angle;
                                node[9] = RingWIn;
                                node[10] = 10;
                                node[11] = 10;
                                node[12] = 300;
                                axon.Add(node);
                                Axons.Add(axon);
                            }
                            sketchPoints++;
                        }
                        skMgr.InsertSketch(true);
                        #endregion
                        if (selMgr.GetSelectedObjectCount() > 0)
                        {
                            selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            modDoc.ClearSelection2(true);
                        }
                        its++;
                    }

                    bools = swModelDocExt.SelectByID2(RefPlanes[its-1].Name, "PLANE", 0, 0, 0, true, 0, null, 0);
                    myRefPlane = featMgr.InsertRefPlane(264, 0.0003, 0, 0, 0, 0);
                    modDoc.ClearSelection2(true);
                    refPlane = (Feature)myRefPlane;
                    RefPlanes.Add(refPlane);
                    bools = swModelDocExt.SelectByID2(refPlane.Name, "PLANE", 0, 0, 0, false, 0, null, 0);
                    skMgr.InsertSketch(false);
                    modDoc.ClearSelection2(true);
                    bools = swModelDocExt.SelectByID2(body.Name, "SOLIDBODY", 0, 0, 0, true, 0, null, 0);
                    modDoc.Sketch3DIntersections();
                    sk = skMgr.ActiveSketch;
                    modDoc.ClearSelection2(true);
                    modDoc.ClearSelection2(true);
                    skMgr.InsertSketch(true);
                    sketchs.Add(sk);
                    bools = swModelDocExt.Rebuild(2);
                    fe = swModelDocExt.GetLastFeatureAdded();
                    sketchFeatures.Add(fe);
                    System.Windows.Forms.MessageBox.Show("A total of ");
                    if (selMgr.GetSelectedObjectCount() > 0)
                    {
                        selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                        modDoc.ClearSelection2(true);
                    }
                    (rotationMatrix, translationMatrix, refPXYZGlobal, refPXYZInSkPlane, verts, vertXYZ, vertexGlobal, vertexInSkPlane, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, vertAngList) = GetIntersectionCurveVertAngles(sk, null, cf, selMgr, modDoc, fe, swModelDocExt);
                    RotationMatricies.Add(rotationMatrix);
                    TranslationMatricies.Add(translationMatrix);
                    RefPtsMatGlobal.Add(refPXYZGlobal);
                    RefPtsMatSketch.Add(refPXYZInSkPlane);
                    AllvertXYZ.Add(vertXYZ);
                    AllvertexGlobal.Add(vertexGlobal);
                    AllvertexInSkPlane.Add(vertexInSkPlane);
                    AllScaledVertsXList.Add(ScaledVertsXList);
                    AllScaledVertsYList.Add(ScaledVertsYList);
                    AllScaledVertsZList.Add(ScaledVertsZList);
                    AllvertAngList.Add(vertAngList);
                    System.Windows.Forms.MessageBox.Show("Centroid of: X: "+refPXYZInSkPlane[0,0]+" Y: " + refPXYZInSkPlane[1, 0] +" Z: " + refPXYZInSkPlane[2, 0]);
                    for (int ap = 0; ap <= (Axons.Count-1); ap++)
                    {
                        double[] intersectionPointsXY = new double[2];
                        double[] distance = new double[10];
                        double angle;
                        double distance2Test;

                        // Generate point for middle third section
                        angle = Axons[ap][0][8];//uni.Sample();
                        (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                        distance2Test = distance[0]*Axons[ap][0][6]; //GenRandFromDistWithBounds(distance[6], distance[3], norm, random);
                        
                            double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                            double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                            SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                            //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                            (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX / cf, dY / cf, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                            double[] node = new double[13];
                            List<double[]> axon = new List<double[]>();
                            node[0] = vPoint[0];
                            node[1] = vPoint[1];
                            node[2] = vPoint[2];
                            node[3] = vPointInSketch[0];
                            node[4] = vPointInSketch[1];
                            node[5] = vPointInSketch[2];
                            node[6] = distance2Test / distance[0];
                            node[7] = distance2Test;
                            node[8] = angle;
                            node[9] = Axons[ap][0][9];
                            node[10] = 10;
                            node[11] = 10;
                            node[12] = 300;
                            Axons[ap].Add(node);

                    }
                    
                    its++;

                }

                //modDoc.ClearSelection2(true);

                //System.Windows.Forms.MessageBox.Show("A total of ");
                //Debug.Print("Arc Count " + sk.GetArcCount());
                //Debug.Print("Edge Contour Count " + sk.GetContourEdgeCount());
                //Debug.Print("Ellipse Count " + sk.GetEllipseCount());
                //Debug.Print("Line Count " + sk.GetLineCount());
                //Debug.Print("Parabola Count " + sk.GetParabolaCount());
                //int pc = 0;
                //Debug.Print("Polyline Count " + sk.GetPolyLineCount(ref pc));
                //Debug.Print("SB Instance Count " + sk.GetSketchBlockInstanceCount());
                //Debug.Print("Contour Count " + sk.GetSketchContourCount());
                //Debug.Print("Path Count " + sk.GetSketchPathCount());
                //Debug.Print("Points Count " + sk.GetSketchPointsCount2());
                //Debug.Print("Region Count " + sk.GetSketchRegionCount());
                //Debug.Print("Slot Count " + sk.GetSketchSlotCount());
                //Debug.Print("Spline Count " + sk.GetSplineCount(ref pc));
                //Debug.Print("Interp Count " + sk.GetSplineInterpolateCount(ref pc));
                //int s;
                //Debug.Print("Parmas Count " + sk.GetSplineParamsCount3(true, out s));
                //Debug.Print("U Points Count " + sk.GetUserPointsCount());

                //object[] swCON = sk.GetSketchContours();
                //SketchContour swSC = (SketchContour)swCON[0];
                //object[] swEdgeO = swSC.GetEdges();
                //Edge swEdge = (Edge)swEdgeO[0];


            }
            else
            {
                System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }
        public (Vertex, double[], Matrix<double>, Matrix<double>, double[]) RotatePointXYZ(double PointXLoc, double PointYLoc, double PointZLoc, object VertexObject, Matrix<double> rotationMatrix, Matrix<double> translationMatrix, bool ToGlobal, bool ToSketch, double cf)
        {
            Vertex myVert = null;
            double[] vPoint = new double[3];
            double[,] pXYZ = new double[3, 1];
            double[] vPointInSketch = new double[3];
            Matrix<double> pXYZGlobal = null;
            Matrix<double> pXYZInSkPlane = null;
            if (Double.IsNaN(PointXLoc))
            {
                myVert = (Vertex)VertexObject;
                vPoint = (double[])myVert.GetPoint();
                double vpX = (double)vPoint[0];
                double vpY = (double)vPoint[1];
                double vpZ = (double)vPoint[2];
                vPoint[0] = vpX * cf;
                vPoint[1] = vpY * cf;
                vPoint[2] = vpZ * cf;
            }
            else if (VertexObject == null)
            {
                vPoint[0] = PointXLoc;
                vPoint[1] = PointYLoc;
                vPoint[2] = PointZLoc;
                myVert = null;
            }

            if (ToSketch)
            {
                pXYZ[0, 0] = vPoint[0];
                pXYZ[1, 0] = vPoint[1];
                pXYZ[2, 0] = vPoint[2];
                pXYZGlobal = Matrix.Create(pXYZ);

                pXYZInSkPlane = Matrix.Multiply(rotationMatrix.Transpose(), pXYZGlobal) + translationMatrix;
                vPointInSketch[0] = pXYZInSkPlane[0, 0];
                vPointInSketch[1] = pXYZInSkPlane[1, 0];
                vPointInSketch[2] = pXYZInSkPlane[2, 0];
            }
            else if (ToGlobal)
            {
                pXYZ[0, 0] = vPoint[0];
                pXYZ[1, 0] = vPoint[1];
                pXYZ[2, 0] = vPoint[2];
                pXYZInSkPlane = Matrix.Create(pXYZ);

                pXYZGlobal = Matrix.Multiply(rotationMatrix, pXYZInSkPlane - translationMatrix);
                vPointInSketch[0] = pXYZInSkPlane[0, 0];
                vPointInSketch[1] = pXYZInSkPlane[1, 0];
                vPointInSketch[2] = pXYZInSkPlane[2, 0];
                vPoint[0] = pXYZGlobal[0, 0];
                vPoint[1] = pXYZGlobal[1, 0];
                vPoint[2] = pXYZGlobal[2, 0];
            }
            

            return (myVert, vPoint, pXYZGlobal, pXYZInSkPlane, vPointInSketch);

        }

        public (bool, int) checkedBand(double distance2Test, double angle, double[] distance, Matrix<double> refPXYZInSkPlane, double cf, bool OuterThird, bool MiddleThird, bool Center, bool OtherRegion)
        {
            bool good = false;
            if (OuterThird)
            {
                if (distance2Test < distance[0] && distance2Test >= distance[3])
                {
                    good = true;
                }
            }
            if (MiddleThird)
            {
                if (distance2Test <= distance[3] && distance2Test >= distance[6])
                {
                    good = true;
                }
            }

            if (Center)
            {
                if (distance2Test <= distance[6] && distance2Test >= 0)
                {
                    good = true;
                }
            }
            int ring = 11;
            double ob;
            double ib;
            if (OtherRegion)
            {
                for (int qe = 0; qe <= 9; qe++)
                {
                    if (qe == 0)
                    {
                        if (distance2Test < distance[0] && distance2Test >= distance[1])
                        {
                            ring = 10;
                        }
                    }
                    else
                    {
                        if (qe == 9)
                        {
                            ob = distance[qe];
                            ib = 0;
                        }
                        else
                        {
                            ob = distance[qe];
                            ib = distance[qe + 1];
                        }
                        
                        if (distance2Test <= ob && distance2Test >= ib)
                        {
                            ring = 10 - qe;
                        }
                    }
                }
            }
            return (good, ring);
        }
        public (double[], double[]) GetIntersection(double angle, List<double> vertAngList, List<double[]> ScaledVertsXList, List<double[]> ScaledVertsYList, List<double[]> ScaledVertsZList, Matrix<double> refPXYZInSkPlane, MLApp.MLApp matlab)
        {
            bool go = true;
            bool found = false;
            int ind2S = 0;
            int a = 0;
            int b = 0;
            object inter = null;
            object[] interPts = new object[2];
            List<double> sortedAngs = new List<double>(vertAngList);
            sortedAngs.Sort();
            while (go)
            {
                if (sortedAngs[ind2S] == angle)
                {
                    found = true;
                }
                else if (ind2S == (sortedAngs.Count-1))
                {
                    a = vertAngList.IndexOf(sortedAngs[ind2S]);
                    b = vertAngList.IndexOf(sortedAngs[0]);
                    found = true;
                }
                else if (angle > sortedAngs[ind2S] && angle < sortedAngs[ind2S + 1])
                {
                    a = vertAngList.IndexOf(sortedAngs[ind2S]);
                    b = vertAngList.IndexOf(sortedAngs[ind2S+1]);
                    found = true;
                }
                //else if (angle > sortedAngs[ind2S])
                //{
                //    if (ind2S == 0)
                //    {
                //        a = vertAngList.Count() - 1;
                //        found = true;
                //    }
                //}

                if (found)
                {
                    double[] compareVectStart = new double[3];
                    compareVectStart[0] = ScaledVertsXList.ElementAt(a)[0];
                    compareVectStart[1] = ScaledVertsYList.ElementAt(a)[0];
                    compareVectStart[2] = ScaledVertsZList.ElementAt(a)[0];
                    double[] compareVectEnd = new double[3];
                    compareVectEnd[0] = ScaledVertsXList.ElementAt(b)[0];
                    compareVectEnd[1] = ScaledVertsYList.ElementAt(b)[0];
                    compareVectEnd[2] = ScaledVertsZList.ElementAt(b)[0];
                    double pt1 = Math.Sqrt(Math.Pow((compareVectEnd[0] - refPXYZInSkPlane[0, 0]), 2) + Math.Pow((compareVectEnd[1] - refPXYZInSkPlane[1, 0]), 2) + Math.Pow((compareVectEnd[2] - refPXYZInSkPlane[2, 0]), 2));
                    double pt2 = Math.Sqrt(Math.Pow((compareVectStart[0] - refPXYZInSkPlane[0, 0]), 2) + Math.Pow((compareVectStart[1] - refPXYZInSkPlane[1, 0]), 2) + Math.Pow((compareVectStart[2] - refPXYZInSkPlane[2, 0]), 2));
                    double maxR = Math.Max(pt1, pt2);

                    double[] refPtXs = new double[2];
                    refPtXs[0] = ScaledVertsXList.ElementAt(a)[0];
                    refPtXs[1] = ScaledVertsXList.ElementAt(b)[0];
                    double[] refPtYs = new double[2];
                    refPtYs[0] = ScaledVertsYList.ElementAt(a)[0];
                    refPtYs[1] = ScaledVertsYList.ElementAt(b)[0];
                    double[] newPtXs = new double[2];
                    newPtXs[0] = refPXYZInSkPlane[0, 0];
                    newPtXs[1] = maxR * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                    double[] newPtYs = new double[2];
                    newPtYs[0] = refPXYZInSkPlane[1, 0];
                    newPtYs[1] = maxR * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                    matlab.Feval("polyxpoly", 2, out inter, refPtXs, refPtYs, newPtXs, newPtYs);
                    interPts = (object[])inter;
                    //double[] intPT = XYZIntersection(compareVectStart, compareVectEnd, useVectStart, useVectEnd);
                    go = false;

                }
                
                ind2S++;
            }
            double[] intPts = new double[2];
            intPts[0] = (double)interPts[0];
            intPts[1] = (double)interPts[1];

            double[] distance = new double[10];
            distance[0] = Math.Sqrt(Math.Pow((double)interPts[0] - refPXYZInSkPlane[0, 0], 2) + Math.Pow((double)interPts[1] - refPXYZInSkPlane[1, 0], 2));
            for (int qw = 1; qw <= 9; qw++)
            {
                double qwd = (double)qw;
                double sc = (1 - (qwd / 10));
                distance[qw] = distance[0] * sc;
            }

            return (intPts, distance);
        }
        public double[] XYZScale(double scale, Vector<double> vect, Vector<double> refvect, Matrix<double> refPt)
        {
            Vector<double> normVect = vect.Clone();
            double mag = normVect.Norm();
            normVect = normVect.Normalize();
            normVect.MultiplyInto(mag * scale, normVect);
            double[] XYZ = new double[3];
            XYZ[0] = normVect.GetValue(0) + refPt[0, 0];
            XYZ[1] = normVect.GetValue(1) + refPt[1, 0];
            XYZ[2] = normVect.GetValue(2) + refPt[2, 0];
            return (XYZ);
        }
       

        public double GenRandFromDistCenter(double maxV, NormalDistribution distributionIN, MersenneTwister randomIN)
        {
            double randStdNormal = distributionIN.Sample(randomIN) / 2.5; //random normal(0,1)
            while (randStdNormal < 0)
            {
                randStdNormal = distributionIN.Sample(randomIN) / 2.5; //random normal(0,1)
            }
            randStdNormal = randStdNormal * maxV;
            return (randStdNormal);
        }
        public double GenRandFromDistWithBounds(double minV, double maxV, NormalDistribution distributionIN, MersenneTwister randomIN)
        {
            double randStdNormal = distributionIN.Sample(randomIN)/2.5; //random normal(0,1)
            randStdNormal = randStdNormal * (maxV-minV)/2;
            randStdNormal = randStdNormal + (maxV + minV) / 2;
            return (randStdNormal);
        }

        public (Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>, List<Vertex>, List<double[]>, List<Matrix<double>>, List<Matrix<double>>, List<double[]>, List<double[]>, List<double[]>, List<double>) GetIntersectionCurveVertAngles(Sketch sk, MathPoint comMathPt, double cf, ISelectionMgr selMgr, IModelDoc2 modDoc, Feature fe, IModelDocExtension swModelDocExt)
        {
            bool bools;

            MathTransform mathTrans = sk.ModelToSketchTransform;
            double[] ad = mathTrans.ArrayData;
            double[,] mT = new double[3, 3];
            double[,] tmT = new double[3, 1];
            mT[0, 0] = ad[0];
            mT[0, 1] = ad[1];
            mT[0, 2] = ad[2];
            mT[1, 0] = ad[3];
            mT[1, 1] = ad[4];
            mT[1, 2] = ad[5];
            mT[2, 0] = ad[6];
            mT[2, 1] = ad[7];
            mT[2, 2] = ad[8];
            tmT[0, 0] = ad[9] * cf;
            tmT[1, 0] = ad[10] * cf;
            tmT[2, 0] = ad[11] * cf;
            Matrix<double> rotationMatrix = Matrix.Create(mT);
            Matrix<double> translationMatrix = Matrix.Create(tmT);
            Matrix<double> refPXYZGlobal = null;
            Matrix<double> refPXYZInSkPlane = null;
            object[] vSkRegions;

            double[,] refPGlobal = new double[3, 1];
            if (comMathPt != null)
            {
                refPGlobal[0, 0] = ((double[])comMathPt.ArrayData)[0] * cf;
                refPGlobal[1, 0] = ((double[])comMathPt.ArrayData)[1] * cf;
                refPGlobal[2, 0] = ((double[])comMathPt.ArrayData)[2] * cf;
                refPXYZGlobal = Matrix.Create(refPGlobal);
                refPXYZInSkPlane = Matrix.Multiply(rotationMatrix.Transpose(), refPXYZGlobal) + translationMatrix;
            }
            else
            {

                Vertex myVertNA;
                double[] vPointNA = null;
                double[] vPointInSketchNA = null;
                bools = swModelDocExt.SelectByID2(fe.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                selMgr.GetSelectedObject5(1);
                object sections = null;
                double[] v1 = swModelDocExt.GetSectionProperties2(sections);
                (myVertNA, vPointNA, refPXYZGlobal, refPXYZInSkPlane, vPointInSketchNA) = RotatePointXYZ(v1[2] * cf, v1[3] * cf, v1[4] * cf, null, rotationMatrix, translationMatrix, false, true, cf);
                if (selMgr.GetSelectedObjectCount() > 0)
                {
                    selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                    modDoc.ClearSelection2(true);
                }

                //bools = swModelDocExt.SelectByID2(fe.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                //modDoc.EditSketch();
                //modDoc.ClearSelection2(true);
                //vSkRegions = (object[])sk.GetSketchRegions();
                //for (int i = 0; i <= (vSkRegions.Count() - 1); i++)
                //{
                //    SketchRegion skRegion = (SketchRegion)vSkRegions[i];
                //    Loop2 myLoop = skRegion.GetFirstLoop();
                //    while (myLoop != null)
                //    {
                //        object[] vVerts = (object[])myLoop.GetVertices();
                //        double[] vXR = new double[vVerts.Count()];
                //        double[] vYR = new double[vVerts.Count()];
                //        double[] vZR = new double[vVerts.Count()];
                //        double[] vXRG = new double[vVerts.Count()];
                //        double[] vYRG = new double[vVerts.Count()];
                //        double[] vZRG = new double[vVerts.Count()];
                //        for (int j = 0; j <= (vVerts.Count() - 1); j++)
                //        {
                //            //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                //            (Vertex myVertTEMP, double[] vPointTEMP, Matrix<double> pXYZGlobalTEMP, Matrix<double> pXYZInSkPlaneTEMP, double[] vPointInSketchTEMP) = RotatePointXYZ(Double.NaN, Double.NaN, Double.NaN, vVerts[j], rotationMatrix, translationMatrix, false, true, cf);

                //            vXR[j] = pXYZInSkPlaneTEMP[0, 0];
                //            vYR[j] = pXYZInSkPlaneTEMP[1, 0];
                //            vZR[j] = pXYZInSkPlaneTEMP[2, 0];

                //            vXRG[j] = pXYZGlobalTEMP[0, 0];
                //            vYRG[j] = pXYZGlobalTEMP[1, 0];
                //            vZRG[j] = pXYZGlobalTEMP[2, 0];
                //        }
                //        Vertex myVertNA;
                //        double[] vPointNA = null;
                //        double[] vPointInSketchNA = null;
                //        (myVertNA, vPointNA, refPXYZGlobal, refPXYZInSkPlane, vPointInSketchNA) = RotatePointXYZ(vXR.Average(), vYR.Average(), vZR.Average(), null, rotationMatrix, translationMatrix, true, false, cf);

                //        myLoop = (Loop2)myLoop.GetNext();
                //    }
                //}
            }
            

            if (selMgr.GetSelectedObjectCount() > 0)
            {
                selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                modDoc.ClearSelection2(true);
            }

            bools = swModelDocExt.SelectByID2(fe.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
            modDoc.EditSketch();
            modDoc.ClearSelection2(true);


            vSkRegions = (object[])sk.GetSketchRegions();
            List<SketchRegion> skRegions = new List<SketchRegion>();
            List<Loop2> myLoops = new List<Loop2>();
            List<Vertex> verts = new List<Vertex>();
            List<double[]> vertXYZ = new List<double[]>();
            List<Matrix<double>> vertexGlobal = new List<Matrix<double>>();
            List<Matrix<double>> vertexInSkPlane = new List<Matrix<double>>();

            List<double> vX = new List<double>();
            List<double> vY = new List<double>();

            double[] ScaledVertsX = new double[10];
            double[] ScaledVertsY = new double[10];
            double[] ScaledVertsZ = new double[10];
            double[] vertAng = new double[10];
            List<double[]> ScaledVertsXList = new List<double[]>();
            List<double[]> ScaledVertsYList = new List<double[]>();
            List<double[]> ScaledVertsZList = new List<double[]>();
            List<double> vertAngList = new List<double>();
            double[] refPt = new double[3];
            refPt[0] = refPXYZInSkPlane[0, 0] - refPXYZInSkPlane[0, 0];
            refPt[1] = refPXYZInSkPlane[1, 0] + 10;
            refPt[2] = refPXYZInSkPlane[2, 0] - refPXYZInSkPlane[2, 0];
            Vector<double> refPtVector = Vector.Create<double>(refPt);

            for (int i = 0; i <= (vSkRegions.Count() - 1); i++)
            {
                SketchRegion skRegion = (SketchRegion)vSkRegions[i];
                skRegions.Add(skRegion);
                Loop2 myLoop = skRegion.GetFirstLoop();
                while (myLoop != null)
                {
                    myLoops.Add(myLoop);
                    object[] vVerts = (object[])myLoop.GetVertices();
                    for (int j = 0; j <= (vVerts.Count() - 1); j++)
                    {
                        //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                        (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(Double.NaN, Double.NaN, Double.NaN, vVerts[j], rotationMatrix, translationMatrix, false, true, cf);
                        verts.Add(myVert);
                        vertXYZ.Add(vPoint);
                        vertexGlobal.Add(pXYZGlobal);
                        vertexInSkPlane.Add(pXYZInSkPlane);

                        double vXR = pXYZInSkPlane[0, 0];
                        double vYR = pXYZInSkPlane[1, 0];
                        double vZR = pXYZInSkPlane[2, 0];
                        vX.Add(vXR);
                        vY.Add(vYR);
                        ScaledVertsX[0] = vXR;
                        ScaledVertsY[0] = vYR;
                        ScaledVertsZ[0] = vZR;

                        double[] angVect = new double[3];
                        angVect[0] = vXR - refPXYZInSkPlane[0, 0];
                        angVect[1] = vYR - refPXYZInSkPlane[1, 0];
                        angVect[2] = vZR - refPXYZInSkPlane[2, 0];
                        Vector<double> angVector = Vector.Create<double>(angVect);
                        double ang = new double();
                        if (angVect[0] > 0)
                        {
                            ang = (Vector.Angle(refPtVector, angVector)) * (180 / Math.PI);
                        }
                        else if (angVect[0] < 0)
                        {
                            ang = 360 - ((Vector.Angle(refPtVector, angVector)) * (180 / Math.PI));
                        }

                        vertAngList.Add(ang);

                        for (int qw = 1; qw <= 9; qw++)
                        {
                            double qwd = (double)qw;
                            double sc = (1 - (qwd / 10));
                            double[] XYZs = XYZScale(sc, angVector, refPtVector, refPXYZInSkPlane);
                            ScaledVertsX[qw] = XYZs[0];
                            ScaledVertsY[qw] = XYZs[1];
                            ScaledVertsZ[qw] = XYZs[2];
                        }
                        double[] tempX = (double[])ScaledVertsX.Clone();
                        double[] tempY = (double[])ScaledVertsY.Clone();
                        double[] tempZ = (double[])ScaledVertsZ.Clone();

                        ScaledVertsXList.Add(tempX);
                        ScaledVertsYList.Add(tempY);
                        ScaledVertsZList.Add(tempZ);

                    }
                    myLoop = (Loop2)myLoop.GetNext();
                }
            }
            return (rotationMatrix, translationMatrix, refPXYZGlobal, refPXYZInSkPlane, verts, vertXYZ, vertexGlobal, vertexInSkPlane, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, vertAngList);
        }

        #region Add Reference Points
        public void AddRefPts()
        {
            // THIS IS A TEST I MADE A CHANGE
            //make sure we have a part open
            // If we are at the beginning of generatong axons then the distance between nodes is 300.5uM, must place the last node 2um before the end of the nerve at the crista
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
                List<Feature> RefPlanes = new List<Feature>();
                List<Feature> sketchFeatures = new List<Feature>();

                List<Feature> centerLinePtsFeat = new List<Feature>();
                List<double[]> centerLinePtsXYZ = new List<double[]>();

                List<Matrix<double>> RotationMatricies = new List<Matrix<double>>();
                List<Matrix<double>> TranslationMatricies = new List<Matrix<double>>();
                List<Sketch> sketchs = new List<Sketch>();
                List<MathPoint> RefMathPt = new List<MathPoint>();
                List<Matrix<double>> RefPtsMatGlobal = new List<Matrix<double>>();
                List<Matrix<double>> RefPtsMatSketch = new List<Matrix<double>>();
                List<List<double[]>> Axons = new List<List<double[]>>();
                List<List<double[]>> AllvertXYZ = new List<List<double[]>>();
                List<List<Matrix<double>>> AllvertexGlobal = new List<List<Matrix<double>>>();
                List<List<Matrix<double>>> AllvertexInSkPlane = new List<List<Matrix<double>>>();
                List<List<double[]>> AllScaledVertsXList = new List<List<double[]>>();
                List<List<double[]>> AllScaledVertsYList = new List<List<double[]>>();
                List<List<double[]>> AllScaledVertsZList = new List<List<double[]>>();
                List<List<double>> AllvertAngList = new List<List<double>>();

                UserUnit unit = (UserUnit)modelDoc.GetUserUnit(0);
                double cf = unit.GetConversionFactor();

                NormalDistribution norm = new NormalDistribution();
                ContinuousUniformDistribution uni = new ContinuousUniformDistribution(0, 360);
                MersenneTwister random = new MersenneTwister();

                MLApp.MLApp matlab = new MLApp.MLApp();

                bool bools;
                Body2 body = default(Body2);
                object myRefPlane = new object();
                Feature refPlane;
                Sketch sk;
                Feature fe;
                Feature combineBody;

                List < MathPoint > refPoints = new List<MathPoint>();
                List<double[]> rpXYZ = new List<double[]>();
                List<Feature> rpFeats = new List<Feature>();
                SketchSegment splineG = null;
                Feature splineFeat = null;
                IBodyFolder bodyFolder;
                List<Feature> bodiesList = new List<Feature>();
                (refPoints, rpXYZ, rpFeats, splineG, splineFeat, centerLinePtsFeat, centerLinePtsXYZ, bodyFolder) = CheckSplinePts((Feature)modelDoc.FirstFeature(), modelDoc);
                if (refPoints == null)
                {
                    (bodyFolder, bodiesList) = GetBodyFolder(feature);
                    Debug.Print(bodyFolder.GetBodyCount().ToString());
                    object[] Bodies = (object[])bodyFolder.GetBodies();



                    //Input: bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr
                    //Output: refPoints, rpXYZ, rpFeats
                    SelectionSet bodySet = GenerateBodySet(bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr);
                    List<double[]> fBoxPts = new List<double[]>();
                    (refPoints, rpXYZ, rpFeats, fBoxPts) = CreatePoints(bodyFolder, Bodies, swModelDocExt, featMgr, modelDoc, selMgr);

                    SelectionSet rpSet = GenerateRPSet(swModelDocExt, featMgr, modelDoc, selMgr, refPoints, rpXYZ, rpFeats);


                    (splineG, splineFeat) = GenSpline(modelDoc, selMgr, skMgr, rpXYZ, cf);

                    System.Windows.Forms.MessageBox.Show("A total of " + refPoints.Count + " reference points were created!");
                    swModelDocExt.Rebuild(2);

                    combineBody = GenCombine(modelDoc, swModelDocExt, selMgr, featMgr, bodiesList);

                    (centerLinePtsFeat, centerLinePtsXYZ) = GenNodePts(selMgr, modelDoc, splineG, cf, rpFeats, rpXYZ, swModelDocExt, splineFeat, featMgr);

                }



                for (int allNode = centerLinePtsFeat.Count - 30; allNode >= centerLinePtsFeat.Count - 35; allNode--)
                {
                    Matrix<double> rotationMatrix = null;
                    Matrix<double> translationMatrix = null;
                    Matrix<double> refPXYZGlobal = null;
                    Matrix<double> refPXYZInSkPlane = null;
                    List<Vertex> verts = new List<Vertex>();
                    List<double[]> vertXYZ = new List<double[]>();
                    List<Matrix<double>> vertexGlobal = new List<Matrix<double>>();
                    List<Matrix<double>> vertexInSkPlane = new List<Matrix<double>>();
                    List<double[]> ScaledVertsXList = new List<double[]>();
                    List<double[]> ScaledVertsYList = new List<double[]>();
                    List<double[]> ScaledVertsZList = new List<double[]>();
                    List<double> vertAngList = new List<double>();
                    if (allNode == centerLinePtsFeat.Count - 30)
                    {
                        Feature firstRefPlane = GenRefPlane(modelDoc, selMgr, swModelDocExt, featMgr, rpFeats, splineG, splineFeat, centerLinePtsFeat[allNode].Name);
                        refPlane = firstRefPlane;
                        RefPlanes.Add(refPlane);

                        (Sketch skT, Feature feT) = GenBoundarySketch(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlane, bodyFolder, skMgr);

                        sketchs.Add(skT);
                        bools = swModelDocExt.Rebuild(2);
                        sketchFeatures.Add(feT);

                        (rotationMatrix, translationMatrix, refPXYZGlobal, refPXYZInSkPlane, verts, vertXYZ, vertexGlobal, vertexInSkPlane, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, vertAngList) = GetIntersectionCurveVertAngles(skT, null, cf, selMgr, modelDoc, feT, swModelDocExt);
                        RotationMatricies.Add(rotationMatrix);
                        TranslationMatricies.Add(translationMatrix);
                        RefPtsMatGlobal.Add(refPXYZGlobal);
                        RefPtsMatSketch.Add(refPXYZInSkPlane);
                        AllvertXYZ.Add(vertXYZ);
                        AllvertexGlobal.Add(vertexGlobal);
                        AllvertexInSkPlane.Add(vertexInSkPlane);
                        AllScaledVertsXList.Add(ScaledVertsXList);
                        AllScaledVertsYList.Add(ScaledVertsYList);
                        AllScaledVertsZList.Add(ScaledVertsZList);
                        AllvertAngList.Add(vertAngList);
                        System.Windows.Forms.MessageBox.Show("Centroid of: X: " + refPXYZInSkPlane[0, 0] + " Y: " + refPXYZInSkPlane[1, 0] + " Z: " + refPXYZInSkPlane[2, 0]);

                        double[] xvs = new double[ScaledVertsXList.Count];
                        double[] yvs = new double[ScaledVertsXList.Count];
                        //matlab.Feval("polyxpoly", 2, out inter, refPtXs, refPtYs, newPtXs, newPtYs);
                        //interPts = (object[])inter;
                        for (int qt = 0; qt <= (ScaledVertsXList.Count - 1); qt++)
                        {
                            xvs[qt] = ScaledVertsXList[qt][0];
                            yvs[qt] = ScaledVertsYList[qt][0];
                        }

                        System.Windows.Forms.MessageBox.Show("Centroid, X: " + xvs.Average() + " Y: " + yvs.Average());
                        bools = swModelDocExt.SelectByID2(feT.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditSketch();

                        int sketchPoints = 1;
                        while (sketchPoints <= 150)
                        {
                            double[] intersectionPointsXY = new double[2];
                            double[] distance = new double[10];
                            double angle;
                            double distance2Test;
                            bool good2Go;
                            int RingWIn;

                            // Generate point for outer third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistWithBounds(distance[3], distance[0], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, true, false, false, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                //object ot;
                                //matlab.Execute(@"cd C:\Users\Brian\Desktop\Solidworks Add Ins");
                                //double rx = refPXYZInSkPlane[0, 0];
                                //double ry = refPXYZInSkPlane[1, 0];
                                //double ipx = intersectionPointsXY[0];
                                //double ipy = intersectionPointsXY[1];
                                //matlab.Feval("PlotPointsAndBound", 1, out ot, xvs, yvs, rx, ry, ipx, ipy, dX, dY);
                                //SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //if (skPoint == null)
                                //{
                                //    System.Windows.Forms.MessageBox.Show("point is null");
                                //}
                                //else
                                //{
                                    //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                    (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX, dY, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                    double[] node = new double[13];
                                    List<double[]> axon = new List<double[]>();
                                    node[0] = vPoint[0];
                                    node[1] = vPoint[1];
                                    node[2] = vPoint[2];
                                    node[3] = vPointInSketch[0];
                                    node[4] = vPointInSketch[1];
                                    node[5] = vPointInSketch[2];
                                    node[6] = distance2Test / distance[0];
                                    node[7] = distance2Test;
                                    node[8] = angle;
                                    node[9] = RingWIn;
                                    node[10] = 10;
                                    node[11] = 10;
                                    node[12] = 300.5;
                                    axon.Add(node);
                                    Axons.Add(axon);
                                //}

                            }

                            // Generate point for middle third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistWithBounds(distance[6], distance[3], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, false, true, false, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                //object ot;
                                //matlab.Execute(@"cd C:\Users\Brian\Desktop\Solidworks Add Ins");
                                //double rx = refPXYZInSkPlane[0, 0];
                                //double ry = refPXYZInSkPlane[1, 0];
                                //double ipx = intersectionPointsXY[0];
                                //double ipy = intersectionPointsXY[1];
                                //matlab.Feval("PlotPointsAndBound", 1, out ot, xvs, yvs, rx, ry, ipx, ipy, dX, dY);
                                //SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //if (skPoint == null)
                                //{
                                //    System.Windows.Forms.MessageBox.Show("point is null");
                                //}
                                //else
                                //{
                                    //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                    (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX, dY, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                    double[] node = new double[13];
                                    List<double[]> axon = new List<double[]>();
                                    node[0] = vPoint[0];
                                    node[1] = vPoint[1];
                                    node[2] = vPoint[2];
                                    node[3] = vPointInSketch[0];
                                    node[4] = vPointInSketch[1];
                                    node[5] = vPointInSketch[2];
                                    node[6] = distance2Test / distance[0];
                                    node[7] = distance2Test;
                                    node[8] = angle;
                                    node[9] = RingWIn;
                                    node[10] = 10;
                                    node[11] = 10;
                                    node[12] = 300.5;
                                    axon.Add(node);
                                    Axons.Add(axon);
                                //}
                            }

                            // Generate point for inner third section
                            angle = uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = GenRandFromDistCenter(distance[6], norm, random);
                            //                                      (generated Distance, generated angle, distance array, ref point, SW conversion factor, TEST OUTER RING, TEST MIDDLE RING, TEST INNER RING, GET SPECIFIC RING INDEX);
                            (good2Go, RingWIn) = checkedBand(distance2Test, angle, distance, refPXYZInSkPlane, cf, false, false, true, true);
                            if (good2Go)
                            {
                                double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                                double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                                //object ot;
                                //matlab.Execute(@"cd C:\Users\Brian\Desktop\Solidworks Add Ins");
                                //double rx = refPXYZInSkPlane[0, 0];
                                //double ry = refPXYZInSkPlane[1, 0];
                                //double ipx = intersectionPointsXY[0];
                                //double ipy = intersectionPointsXY[1];
                                //matlab.Feval("PlotPointsAndBound", 1, out ot, xvs, yvs, rx, ry, ipx, ipy, dX, dY);
                                //SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                                //if (skPoint == null)
                                //{
                                //    System.Windows.Forms.MessageBox.Show("point is null");
                                //}
                                //else
                                //{
                                    //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                                    (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX, dY, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                                    double[] node = new double[13];
                                    List<double[]> axon = new List<double[]>();
                                    node[0] = vPoint[0];
                                    node[1] = vPoint[1];
                                    node[2] = vPoint[2];
                                    node[3] = vPointInSketch[0];
                                    node[4] = vPointInSketch[1];
                                    node[5] = vPointInSketch[2];
                                    node[6] = distance2Test / distance[0];
                                    node[7] = distance2Test;
                                    node[8] = angle;
                                    node[9] = RingWIn;
                                    node[10] = 10;
                                    node[11] = 10;
                                    node[12] = 300.5;
                                    axon.Add(node);
                                    Axons.Add(axon);
                                //}
                            }
                            sketchPoints++;
                        }
                        double[] dx = new double[(Axons.Count() - 1)];
                        double[] dy = new double[(Axons.Count() - 1)];
                        for (int allA = 0; allA <= (Axons.Count()-2); allA++)
                        {
                            dx[allA] = Axons[allA][0][3];
                            dy[allA] = Axons[allA][0][4];
                            //if (selMgr.GetSelectedObjectCount() > 0)
                            //{
                            //    selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            //    modelDoc.ClearSelection2(true);
                            //}
                            //SketchPoint skPoint = skMgr.CreatePoint(Axons[allA][0][3] / cf, Axons[allA][0][4] / cf, 0);
                            //if (skPoint == null)
                            //{
                            //    System.Windows.Forms.MessageBox.Show("point is null"+Axons[allA][0][9]);
                            //}

                        }
                        object ot;
                        matlab.Execute(@"cd C:\Users\Brian\Desktop\Solidworks Add Ins");
                        double rx = refPXYZInSkPlane[0, 0];
                        double ry = refPXYZInSkPlane[1, 0];
                        matlab.Feval("PlotPointsAndBound", 1, out ot, xvs, yvs, rx, ry, null, null, dx, dy);
                        skMgr.InsertSketch(true);

                        bools = swModelDocExt.SelectByID2(feT.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditDelete();
                        bools = swModelDocExt.SelectByID2(refPlane.Name, "PLANE", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditDelete();

                        if (selMgr.GetSelectedObjectCount() > 0)
                        {
                            selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            modelDoc.ClearSelection2(true);
                        }

                    }
                    else
                    {
                        Feature firstRefPlane = GenRefPlane(modelDoc, selMgr, swModelDocExt, featMgr, rpFeats, splineG, splineFeat, centerLinePtsFeat[allNode].Name);
                        refPlane = firstRefPlane;
                        RefPlanes.Add(refPlane);

                        (Sketch skT, Feature feT) = GenBoundarySketch(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlane, bodyFolder, skMgr);

                        sketchs.Add(skT);
                        bools = swModelDocExt.Rebuild(2);
                        sketchFeatures.Add(feT);

                        (rotationMatrix, translationMatrix, refPXYZGlobal, refPXYZInSkPlane, verts, vertXYZ, vertexGlobal, vertexInSkPlane, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, vertAngList) = GetIntersectionCurveVertAngles(skT, null, cf, selMgr, modelDoc, feT, swModelDocExt);
                        RotationMatricies.Add(rotationMatrix);
                        TranslationMatricies.Add(translationMatrix);
                        RefPtsMatGlobal.Add(refPXYZGlobal);
                        RefPtsMatSketch.Add(refPXYZInSkPlane);
                        AllvertXYZ.Add(vertXYZ);
                        AllvertexGlobal.Add(vertexGlobal);
                        AllvertexInSkPlane.Add(vertexInSkPlane);
                        AllScaledVertsXList.Add(ScaledVertsXList);
                        AllScaledVertsYList.Add(ScaledVertsYList);
                        AllScaledVertsZList.Add(ScaledVertsZList);
                        AllvertAngList.Add(vertAngList);
                        //System.Windows.Forms.MessageBox.Show("Centroid of: X: " + refPXYZInSkPlane[0, 0] + " Y: " + refPXYZInSkPlane[1, 0] + " Z: " + refPXYZInSkPlane[2, 0]);

                        double[] xvs = new double[ScaledVertsXList.Count];
                        double[] yvs = new double[ScaledVertsXList.Count];
                        //matlab.Feval("polyxpoly", 2, out inter, refPtXs, refPtYs, newPtXs, newPtYs);
                        //interPts = (object[])inter;
                        for (int qt = 0; qt <= (ScaledVertsXList.Count - 1); qt++)
                        {
                            xvs[qt] = ScaledVertsXList[qt][0];
                            yvs[qt] = ScaledVertsYList[qt][0];
                        }

                        //System.Windows.Forms.MessageBox.Show("Centroid, X: " + xvs.Average() + " Y: " + yvs.Average());
                        bools = swModelDocExt.SelectByID2(feT.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditSketch();

                        for (int ap = 0; ap <= (Axons.Count - 1); ap++)
                        {
                            double[] intersectionPointsXY = new double[2];
                            double[] distance = new double[10];
                            double angle;
                            double distance2Test;

                            // Generate point for middle third section
                            angle = Axons[ap][0][8];//uni.Sample();
                            (intersectionPointsXY, distance) = GetIntersection(angle, vertAngList, ScaledVertsXList, ScaledVertsYList, ScaledVertsZList, refPXYZInSkPlane, matlab);

                            distance2Test = distance[0] * Axons[ap][0][6]; //GenRandFromDistWithBounds(distance[6], distance[3], norm, random);

                            double dX = distance2Test * Math.Sin(angle * (Math.PI / 180)) + refPXYZInSkPlane[0, 0];
                            double dY = distance2Test * Math.Cos(angle * (Math.PI / 180)) + refPXYZInSkPlane[1, 0];
                            //SketchPoint skPoint = skMgr.CreatePoint(dX / cf, dY / cf, 0);
                            //Vertex obj,   vertex XYZ,       point in global,          point in sketch,             vertex in sketch,                          X or NaN,   Y or NaN,   Z or NaN, vertex obj or null, rotation matrix, rotate to global, rotate to sketch, conversion factor
                            (Vertex myVert, double[] vPoint, Matrix<double> pXYZGlobal, Matrix<double> pXYZInSkPlane, double[] vPointInSketch) = RotatePointXYZ(dX, dY, 0, null, rotationMatrix, translationMatrix, true, false, cf);
                            double[] node = new double[13];
                            List<double[]> axon = new List<double[]>();
                            node[0] = vPoint[0];
                            node[1] = vPoint[1];
                            node[2] = vPoint[2];
                            node[3] = vPointInSketch[0];
                            node[4] = vPointInSketch[1];
                            node[5] = vPointInSketch[2];
                            node[6] = distance2Test / distance[0];
                            node[7] = distance2Test;
                            node[8] = angle;
                            node[9] = Axons[ap][0][9];
                            node[10] = 10;
                            node[11] = 10;
                            node[12] = 300;
                            Axons[ap].Add(node);

                        }
                        double[] dx = new double[(Axons.Count() - 1)];
                        double[] dy = new double[(Axons.Count() - 1)];
                        for (int allA = 0; allA <= (Axons.Count() - 2); allA++)
                        {
                            dx[allA] = Axons[allA][0][3];
                            dy[allA] = Axons[allA][0][4];
                            //if (selMgr.GetSelectedObjectCount() > 0)
                            //{
                            //    selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            //    modelDoc.ClearSelection2(true);
                            //}
                            //SketchPoint skPoint = skMgr.CreatePoint(Axons[allA][0][3] / cf, Axons[allA][0][4] / cf, 0);
                            //if (skPoint == null)
                            //{
                            //    System.Windows.Forms.MessageBox.Show("point is null"+Axons[allA][0][9]);
                            //}

                        }
                        object ot;
                        //matlab.Execute(@"cd C:\Users\Brian\Desktop\Solidworks Add Ins");
                        double rx = refPXYZInSkPlane[0, 0];
                        double ry = refPXYZInSkPlane[1, 0];
                        //matlab.Feval("PlotPointsAndBound", 1, out ot, xvs, yvs, rx, ry, null, null, dx, dy);
                        skMgr.InsertSketch(true);

                        bools = swModelDocExt.SelectByID2(feT.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditDelete();
                        bools = swModelDocExt.SelectByID2(refPlane.Name, "PLANE", 0, 0, 0, false, 0, null, 0);
                        modelDoc.EditDelete();

                        if (selMgr.GetSelectedObjectCount() > 0)
                        {
                            selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                            modelDoc.ClearSelection2(true);
                        }

                    }
                }


                //List<Feature> coordSystems = GenRefAxis(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlanesFeat);

                //Feature combineFeat = GenCombine(modelDoc, swModelDocExt, selMgr, featMgr, bodiesList);

                //List<Sketch> boundryCurves = GenBoundarySketches(selMgr, modelDoc, swModelDocExt, featMgr, rpFeats, refPlanesFeat, bodyFolder, fBoxPts, skMgr);

                if (selMgr.GetSelectedObjectCount() > 0)
                {
                    selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                    modelDoc.ClearSelection2(true);
                }

                List<SketchSegment> axonSpline = new List<SketchSegment>();
                skMgr.Insert3DSketch(true);
                Sketch Axnsk = skMgr.ActiveSketch;
                Feature splineSketchFeat = (Feature)Axnsk;
                DataTable table = new DataTable();
                DataColumn column;
                DataRow row;
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Axon Name";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Sketch Name";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Int32");
                column.ColumnName = "Node Number";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Int32");
                column.ColumnName = "Ring Number";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Axon Diameter";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Axon Type";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Distance to Next Node";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Units";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Global X";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Global Y";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Global Z";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Sketch X";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Sketch Y";
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Node Sketch Z";
                table.Columns.Add(column);

                for (int axnNum = 0; axnNum <= (Axons.Count-2); axnNum++)
                {
                    if (selMgr.GetSelectedObjectCount() > 0)
                    {
                        selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                        modelDoc.ClearSelection2(true);
                    }
                    double[] Apoints = new double[(Axons[axnNum].Count() - 1) * 3];
                    object ApointArray = null;
                    for (int aNum = 0; aNum <= (Axons[axnNum].Count()-2); aNum++)
                    {
                        Apoints[aNum + 2 * aNum] = Axons[axnNum][aNum][0] / cf;
                        Apoints[aNum + 1 + 2 * aNum] = Axons[axnNum][aNum][1] / cf;
                        Apoints[aNum + 2 + 2 * aNum] = Axons[axnNum][aNum][2] / cf;

                    }
                    ApointArray = Apoints;
                    
                    SketchSegment swSketchSegment = (SketchSegment)skMgr.CreateSpline2((ApointArray), true);
                    axonSpline.Add(swSketchSegment);
                    modelDoc.ClearSelection2(true);

                    for (int aNum = 0; aNum <= (Axons[axnNum].Count() - 2); aNum++)
                    {
                        row = table.NewRow();
                        row["Axon Name"] = swSketchSegment.GetName();
                        row["Sketch Name"] = splineSketchFeat.Name;
                        row["Node Number"] = aNum + 1;
                        row["Ring Number"] = Axons[axnNum][aNum][9];
                        row["Axon Diameter"] = Axons[axnNum][aNum][10];
                        row["Axon Type"] = Axons[axnNum][aNum][11];
                        row["Distance to Next Node"] = Axons[axnNum][aNum][12];
                        row["Units"] = "meters";
                        row["Node Global X"] = Axons[axnNum][aNum][0] / cf;
                        row["Node Global Y"] = Axons[axnNum][aNum][1] / cf;
                        row["Node Global Z"] = Axons[axnNum][aNum][2] / cf;
                        row["Node Sketch X"] = Axons[axnNum][aNum][3] / cf;
                        row["Node Sketch Y"] = Axons[axnNum][aNum][4] / cf;
                        row["Node Sketch Z"] = Axons[axnNum][aNum][5] / cf;
                        table.Rows.Add(row);
                    }


                }
                skMgr.Insert3DSketch(true);
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                SaveFileDialog1.Title = "Choose Where to Save the .CSV File";
                SaveFileDialog1.DefaultExt = "csv";
                SaveFileDialog1.AddExtension = true;
                SaveFileDialog1.Filter = "CSV Files (*.csv)|*.csv";
                SaveFileDialog1.ShowDialog();
                //String PathName = modelDoc.GetPathName();
                table.ToCSV(SaveFileDialog1.FileName);

            }
            else
            {

            }

                

        }

        public (Sketch, Feature) GenBoundarySketch(ISelectionMgr selMgrIN, IModelDoc2 modelDocIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, List<Feature> rpFeatsIN, Feature refPlanesFeatIN, IBodyFolder bodyFolderIN, SketchManager skMgrIN)
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
            Sketch crossSections = null;
            //Debug.Print("       Name of feature: " + (string)combineSB.Name);

            bool status;
            status = swModelDocExtIN.SelectByID2(refPlanesFeatIN.Name, "PLANE", 0, 0, 0, false, 0, null, 0);
            skMgrIN.InsertSketch(true);
            modelDocIN.ClearSelection2(true);
            modelDocIN.Sketch3DIntersections();
            status = swModelDocExtIN.SelectByID2(combineSB.Name, "SOLIDBODY", 0, 0, 0, true, 0, null, 0);
            modelDocIN.Sketch3DIntersections();
            crossSections = skMgrIN.ActiveSketch;
            modelDocIN.ClearSelection2(true);
            modelDocIN.ClearSelection2(true);
            skMgrIN.InsertSketch(true);
            Feature fe = swModelDocExtIN.GetLastFeatureAdded();
            return (crossSections, fe);
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

        public Feature GenRefPlane(IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, List<Feature> rpFeatsIN, SketchSegment splineGIN, Feature splineFeatIN, String pointName)
        {
            if (selMgrIN.GetSelectedObjectCount() > 0)
            {
                selMgrIN.DeSelect2(Enumerable.Range(1, selMgrIN.GetSelectedObjectCount()).ToArray(), -1);
                modelDocIN.ClearSelection2(true);
            }
            bool bools;
            //List<Feature> refPlanes = new List<Feature>();
            Feature refPlane;
            //foreach (IFeature fs in rpFeatsIN)
            //{
                bools = swModelDocExtIN.SelectByID2(splineGIN.GetName() + "@" + splineFeatIN.Name, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //bools = swModelDocExtIN.SelectByID2(fs.Name, "DATUMPOINT", 0, 0, 0, true, 1, null, 0);
            if (pointName == null)
            {
                bools = swModelDocExtIN.SelectByID2(rpFeatsIN[1].Name, "DATUMPOINT", 0, 0, 0, true, 1, null, 0);
            }
            else
            {
                bools = swModelDocExtIN.SelectByID2(pointName, "DATUMPOINT", 0, 0, 0, true, 1, null, 0);
            }
            
            //Debug.Print("Num of selected" + selMgrIN.GetSelectedObjectCount());
                object myRefPlane = featMgrIN.InsertRefPlane(2, 0, 4, 0, 0, 0);
            //Feature refPlane = (Feature)myRefPlane;
            //refPlanes.Add(refPlane);
            refPlane = (Feature)myRefPlane;
                modelDocIN.ClearSelection2(true);
            //}
            return (refPlane);
        }

        public (List<Feature>, List<double[]>) GenNodePts(ISelectionMgr selMgr, IModelDoc2 modelDoc, SketchSegment splineG, double cf, List<Feature> rpFeats, List<double[]> rpXYZ, IModelDocExtension swModelDocExt, Feature splineFeat, IFeatureManager featMgr)
        {
            if (selMgr.GetSelectedObjectCount() > 0)
            {
                selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                modelDoc.ClearSelection2(true);
            }
            List<Feature> clpf = new List<Feature>();
            List<double[]> clpXYZ = new List<double[]>();
            bool bools;
            double nStartParam;
            double nEndParam;
            double newLength = 0;
            bool bIsClosed;
            bool bIsPeriodic;
            int ii = 0;
            List<double[]> nodePoints = new List<double[]>();
            List<SketchPoint> sketchPoints = new List<SketchPoint>();
            Curve splineCurve = (Curve)splineG.GetCurve();
            splineCurve.GetEndParams(out nStartParam, out nEndParam, out bIsClosed, out bIsPeriodic);
            double curveL = splineCurve.GetLength3(nStartParam, nEndParam) * cf * 1000;
            newLength = curveL;
            double percent = 0;
            while (newLength > 0)
            {
                if (ii == 0)
                {
                    clpf.Add(rpFeats.Last());
                    clpXYZ.Add(rpXYZ.Last());
                    ii++;
                }
                else if (ii == 1)
                {
                    newLength = newLength - 1;
                    percent = (newLength / curveL) * 100;
                    bools = swModelDocExt.SelectByID2(splineG.GetName() + "@" + splineFeat.Name, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                    object[] c = featMgr.InsertReferencePoint(2, 1, percent, 1);
                    object c2 = (object)c[0];
                    Feature cF = (Feature)c2;
                    RefPoint rp = (RefPoint)cF.GetSpecificFeature2();
                    MathPoint mp = rp.GetRefPoint();
                    double[] mp2 = mp.ArrayData;
                    clpf.Add(cF);
                    clpXYZ.Add(mp2);
                    if (selMgr.GetSelectedObjectCount() > 0)
                    {
                        selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                        modelDoc.ClearSelection2(true);
                    }
                    ii++;
                }
                else if (ii == 2)
                {
                    newLength = newLength - 310;
                    percent = (newLength / curveL) * 100;
                    bools = swModelDocExt.SelectByID2(splineG.GetName() + "@" + splineFeat.Name, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                    object[] c = (object[])featMgr.InsertReferencePoint(2, 1, percent, 1);
                    object c2 = (object)c[0];
                    Feature cF = (Feature)c2;
                    RefPoint rp = (RefPoint)cF.GetSpecificFeature2();
                    MathPoint mp = rp.GetRefPoint();
                    double[] mp2 = mp.ArrayData;
                    clpf.Add(cF);
                    clpXYZ.Add(mp2);
                    if (selMgr.GetSelectedObjectCount() > 0)
                    {
                        selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                        modelDoc.ClearSelection2(true);
                    }
                    ii++;
                }
                else
                {
                    newLength = newLength - 300.5;
                    percent = (newLength / curveL) * 100;
                    bools = swModelDocExt.SelectByID2(splineG.GetName() + "@" + splineFeat.Name, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                    object[] c = (object[])featMgr.InsertReferencePoint(2, 1, percent, 1);
                    object c2 = (object)c[0];
                    Feature cF = (Feature)c2;
                    RefPoint rp = (RefPoint)cF.GetSpecificFeature2();
                    MathPoint mp = rp.GetRefPoint();
                    double[] mp2 = mp.ArrayData;
                    clpf.Add(cF);
                    clpXYZ.Add(mp2);
                    if (selMgr.GetSelectedObjectCount() > 0)
                    {
                        selMgr.DeSelect2(Enumerable.Range(1, selMgr.GetSelectedObjectCount()).ToArray(), -1);
                        modelDoc.ClearSelection2(true);
                    }
                    ii++;
                }
                if ((newLength - 300.5) < 0)
                {
                    newLength = -100;
                }
            }
            return (clpf, clpXYZ);
        }
        public (SketchSegment, Feature) GenSpline(IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, SketchManager skMgrIN, List<double[]> rpXYZIN, double cf)
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

        public SelectionSet GenerateRPSet(IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, IModelDoc2 modelDocIN, ISelectionMgr selMgrIN, List<MathPoint> refPointsIN, List<double[]> rpXYZIN, List<Feature> rpFeatsIN)
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

        public (List<MathPoint> refPointsOUT, List<double[]> rpXYZOUT, List<Feature> rpFeatsOUT, List<double[]> faceBox) CreatePoints(IBodyFolder bodyFolderIN, object[] BodiesIN, IModelDocExtension swModelDocExtIN, IFeatureManager featMgrIN, IModelDoc2 modelDocIN, ISelectionMgr selMgrIN)
        {
            bool boolstatus;
            bool addFlag = true;
            List<double[]> faceD = new List<double[]>(); 
            List<double[]> faceBox = new List<double[]>();
            for (int i = 0; i <= (bodyFolderIN.GetBodyCount() - 3); i++)// Do not include the last body, use self placed point instead
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
                                else if (i == (bodyFolderIN.GetBodyCount() - 2))
                                {
                                    if ((faceD.Count % (bodyFolderIN.GetBodyCount() - 1)) == 1)
                                    {
                                        addFlag = false;
                                        k = (faceD.Count - 1);
                                    }
                                    else
                                    {
                                        addFlag = true;
                                    }

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

            Feature feature2 = (Feature)modelDocIN.FirstFeature();
            bool firstPass = true;
            List<MathPoint> refPointsOUT = new List<MathPoint>();
            List<double[]> rpXYZOUT = new List<double[]>();
            List<Feature> rpFeatsOUT = new List<Feature>();

            MathPoint rpHold = null;
            double[] rpXYZHold = new double[3];
            Feature rpFeatsHold = null;

            int refptCount = 0;
            while (feature2 != null)
            {
                string FeatType = feature2.Name;
                string FeatTypeName = feature2.GetTypeName2();
                //Debug.Print(" " + FeatType + " [" + FeatTypeName + "]");
                if (FeatTypeName == "RefPoint")
                {
                    if (refptCount == 0)
                    {
                        RefPoint rp = (RefPoint)feature2.GetSpecificFeature2();
                        rpHold = rp.GetRefPoint();
                        rpXYZHold = rpHold.ArrayData;
                        rpFeatsHold = feature2;
                    }
                    else
                    {
                        RefPoint rp = (RefPoint)feature2.GetSpecificFeature2();
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
                    refptCount++;
                    
                }
                feature2 = (Feature)feature2.GetNextFeature();
                if ((feature2 == null) && (firstPass == true))
                {
                    modelDocIN.EditDelete();
                    feature2 = (Feature)modelDocIN.FirstFeature();
                    firstPass = false;
                    refPointsOUT.Clear();
                    rpXYZOUT.Clear();
                    rpFeatsOUT.Clear();
                    refptCount = 0;
                }
                else if ((feature2 == null) && (firstPass == false))
                {
                    refPointsOUT.Add(rpHold);
                    rpXYZOUT.Add(rpXYZHold);
                    rpFeatsOUT.Add(rpFeatsHold);
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

        public (List<MathPoint>, List<double[]>, List<Feature>, SketchSegment, Feature, List<Feature>, List<double[]>, IBodyFolder) CheckSplinePts(Feature swFeature, IModelDoc2 modelDoc)
        {
            bool hasPts = false;
            List<MathPoint> refPoints = new List<MathPoint>();
            List<double[]> rpXYZ = new List<double[]>();
            List<Feature> rpFeats = new List<Feature>();
            SketchSegment splineG = null;
            Feature splineFeat = null;
            List<Feature> centerLinePtsFeat = new List<Feature>();
            List<double[]> centerLinePtsXYZ = new List<double[]>();
            IBodyFolder swBodyFolder = null;
            bool gotOne = false;
            while (swFeature != null)
            {
                string FeatType = swFeature.Name;
                string FeatTypeName = swFeature.GetTypeName2();
                if (FeatTypeName == "RefPoint" && FeatType == "Point1")
                {
                    gotOne = true;
                }
                else if (FeatTypeName == "RefPoint" && gotOne == true)
                {
                    hasPts = true;
                    swFeature = null;
                }
                if (!hasPts)
                {
                    swFeature = (Feature)swFeature.GetNextFeature();
                }
                

            }

            if (hasPts)
            {
                swFeature = modelDoc.FirstFeature();
                bool sk3D = false;
                while (swFeature != null)
                {
                    string FeatType = swFeature.Name;
                    string FeatTypeName = swFeature.GetTypeName2();
                    if (FeatTypeName == "SolidBodyFolder")
                    {
                        swBodyFolder = (IBodyFolder)swFeature.GetSpecificFeature2();
                    }
                    else if (FeatTypeName == "RefPoint" && sk3D == false)
                    {
                        Feature tempFeat = swFeature;
                        RefPoint rp = (RefPoint)tempFeat.GetSpecificFeature2();
                        MathPoint mp = rp.GetRefPoint();
                        double[] mp2 = mp.ArrayData;
                        refPoints.Add(mp);
                        rpXYZ.Add(mp2);
                        rpFeats.Add(tempFeat);
                    }
                    else if (FeatType == "3DSketch1")
                    {
                        sk3D = true;
                        Feature tempFeat = swFeature;
                        Sketch sk = (Sketch)tempFeat.GetSpecificFeature2();
                        object[] sso = sk.GetSketchSegments();
                        splineG = (SketchSegment)sso[0];
                        splineFeat = tempFeat;
                        centerLinePtsFeat.Add(rpFeats.Last());
                        centerLinePtsXYZ.Add(rpXYZ.Last());
                    }
                    else if (FeatTypeName == "RefPoint" && sk3D == true)
                    {
                        Feature tempFeat = swFeature;
                        RefPoint rp = (RefPoint)tempFeat.GetSpecificFeature2();
                        MathPoint mp = rp.GetRefPoint();
                        double[] mp2 = mp.ArrayData;
                        centerLinePtsFeat.Add(tempFeat);
                        centerLinePtsXYZ.Add(mp2);
                    }
                    swFeature = swFeature.GetNextFeature();
                }
                return (refPoints, rpXYZ, rpFeats, splineG, splineFeat, centerLinePtsFeat, centerLinePtsXYZ, swBodyFolder);
            }
            else
            {
                return (null, null, null, null, null, null, null, null);
            }
            
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
