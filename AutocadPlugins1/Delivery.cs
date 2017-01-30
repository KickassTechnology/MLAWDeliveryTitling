using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Publishing;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;
using System.Threading;

namespace AutocadPlugins1
{
    public partial class Delivery : Form
    {
        public Delivery()
        {
            InitializeComponent();
            DateTime dtTime = DateTime.Now;
            txtDate.Text = dtTime.ToString("MM-dd-yyyy");
        }

        public static String strJobNumber;
        public static String strFileName;
        public static String strDeliveryFile;
        public static String strAddress;

        public static double dSeal1PosX;
        public static double dSeal1PosY;

        public static double dSeal2PosX;
        public static double dSeal2PosY;

        public static double dSeal3PosX;
        public static double dSeal3PosY;

        private void button1_Click(object sender, EventArgs e)
        {
            String strJobNumber = textBox1.Text;
            String strFormat = "DWG";

            if (radioButton1.Checked == true)
            {
                strFormat = "DWF";
            }

            if (radioButton3.Checked == true)
            {
                strFormat = "PDF";
            }

            lstSelected.Items.Add(strFormat + " - " + strJobNumber);
            textBox1.Text = "";

        }

        private void deliverJobs()
        {
            foreach (string s in lstSelected.Items)
            {
                //get our input from the list
                String[] strItem = s.Split('-');
                String strFormat = strItem[0].Trim();
                strJobNumber = strItem[1].Trim();

                //To find the file on the server, we need to use the date from the job number, not the current date.
                String strCurYear = "20" + strJobNumber.Substring(0, 2);

                //Path to the fileserver on
                string drive = Path.GetPathRoot("M:");   
                String strPath = "";
                string strImgPath = "";

                if (!Directory.Exists(drive))
                {
                    //If we're testing on my PC 
                    strPath = "C:\\Users\\Joe\\Downloads\\";
                    strImgPath = "C:\\Users\\Joe\\Downloads\\";
                }
                else
                {
                    //Live MLAW environment
                    strPath = "M:\\SFR Foundations\\" + strCurYear + " Final Designs\\";
                    strImgPath = "M:\\Toolbox\\Support 2016\\Seals\\";
                }

                //Path to the file
                strFileName = strPath + strJobNumber + ".dwg";
                DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                //Date to use for the sealing
                String strDate = txtDate.Text;

                if (File.Exists(strFileName))
                {
                    //get the dwg file
                    acDocMgr.Open(strFileName, false);

                    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Database db = doc.Database;

                    //if this customer wants the larger sized delivable
                    if (isLargeSize() == false)
                    {
                        dSeal1PosX = 1400.0;
                        dSeal1PosY = 750.0;
                    }
                    else
                    {

                        using (Transaction acTrans = doc.Database.TransactionManager.StartTransaction())
                        {
                            doc.LockDocument();
                            LayerTable acLyrTbl;
                            acLyrTbl = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                            if (radioButton5.Checked == true)
                            {
                                LayerTableRecord acLyrTblRec = acTrans.GetObject(acLyrTbl["9_SEAL MRL"], OpenMode.ForRead) as LayerTableRecord;
                                acLyrTblRec.UpgradeOpen();
                                acLyrTblRec.IsFrozen = false;
                            }
                            else if (radioButton6.Checked == true)
                            {
                                LayerTableRecord acLyrTblRec = acTrans.GetObject(acLyrTbl["9_SEAL EJU"], OpenMode.ForRead) as LayerTableRecord;
                                acLyrTblRec.UpgradeOpen();
                                acLyrTblRec.IsFrozen = false;
                            }
                            else
                            {
                                LayerTableRecord acLyrTblRec = acTrans.GetObject(acLyrTbl["9_SEAL CBC"], OpenMode.ForRead) as LayerTableRecord;
                                acLyrTblRec.UpgradeOpen();
                                acLyrTblRec.IsFrozen = false;
                            }
                           
                            acTrans.Commit();
                        }
                        
                    }

                    using (Transaction newTransaction = doc.Database.TransactionManager.StartTransaction())
                    {
                        //open up the doc
                        doc.LockDocument();
                        BlockTable newBlockTable;
                        newBlockTable = newTransaction.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord newBlockTableRecord;
                        newBlockTableRecord = (BlockTableRecord)newTransaction.GetObject(newBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        TextStyleTable newTextStyleTable = newTransaction.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                        // Define the name and image to use
                        string strImgName = "Seal";
                        string strImgName2 = "Seal2";
                        string strImgName3 = "Seal3";


                        //get the image file that we need
                        string strImgFileName = "";
                        
                        //all of the images used for seals are different formats and different sizes. The resize fixes that.
                        Double imgResize;
                        
                        if (radioButton5.Checked == true)
                        {
                            strImgFileName = strImgPath + "Mike.tif";
                            imgResize = 0.7;
                        }
                        else if (radioButton6.Checked == true)
                        {
                            strImgFileName = strImgPath + "Eric.bmp";
                            imgResize = 4.4;
                        }
                        else
                        {
                            strImgFileName = strImgPath + "Christine.png";
                            imgResize = 0.8;
                        }

                        //create the objects we need
                        bool bRasterDefCreated = false;

                        //seal goes in 2 places for some drawing, 3 for others
                        ObjectId acImgDefId = new ObjectId();
                        ObjectId acImgDefId2 = new ObjectId();
                        ObjectId acImgDefId3 = new ObjectId();

                        ObjectId acImgDctID = RasterImageDef.GetImageDictionary(db);

                        // Check to see if the dictionary does not exist, if not then create it
                        if (acImgDctID.IsNull)
                        {
                            acImgDctID = RasterImageDef.CreateImageDictionary(db);
                        }

                        // Open the image dictionary
                        DBDictionary acImgDict = newTransaction.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;

                        // Create a raster image definition
                        RasterImageDef acRasterDefNew = new RasterImageDef();
                        RasterImageDef acRasterDefNew2 = new RasterImageDef();
                        RasterImageDef acRasterDefNew3 = new RasterImageDef();


                        // Set the source for the image file
                        acRasterDefNew.SourceFileName = strImgFileName;
                        acRasterDefNew2.SourceFileName = strImgFileName;
                        acRasterDefNew3.SourceFileName = strImgFileName;


                        // Load the image into memory
                        try
                        {
                            acRasterDefNew.Load();
                            acRasterDefNew2.Load();
                            acRasterDefNew3.Load();
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show("Could not load the seal:" + ex.Message);
                        }

                    

                        // Add the image definition to the dictionary
                        acImgDict.UpgradeOpen();
                        acImgDefId = acImgDict.SetAt(strImgName, acRasterDefNew);
                        acImgDefId2 = acImgDict.SetAt(strImgName2, acRasterDefNew2);
                        acImgDefId3 = acImgDict.SetAt(strImgName3, acRasterDefNew3);

                        newTransaction.AddNewlyCreatedDBObject(acRasterDefNew, true);
                        newTransaction.AddNewlyCreatedDBObject(acRasterDefNew2, true);
                        newTransaction.AddNewlyCreatedDBObject(acRasterDefNew3, true);

                        bRasterDefCreated = true;


                        BlockTable acBlkTbl;
                        acBlkTbl = newTransaction.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = newTransaction.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        //Get the address
                        foreach (var blkId in acBlkTblRec)
                        {
                            var acBlock = newTransaction.GetObject(blkId, OpenMode.ForRead) as BlockReference;
                            if (acBlock == null) continue;

                            //if it's a regular sized drawing, get the address from the title block
                            if (isLargeSize() == false)
                            {
                                if (!acBlock.Name.Equals("TITLE_BLOCK", StringComparison.CurrentCultureIgnoreCase)) continue;
                                foreach (ObjectId attId in acBlock.AttributeCollection)
                                {
                                    var acAtt = newTransaction.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                    if (acAtt == null) continue;

                                    if (!acAtt.Tag.Equals("ADDRESS", StringComparison.CurrentCultureIgnoreCase)) continue;
                                    acAtt.UpgradeOpen();
                                    strAddress = acAtt.TextString;
                                }
                            }
                            else
                            {
                                //if we have a 24x36, the address is not in the address attribute. It's in the Project name attribute
                                if (!acBlock.Name.Equals("T-block 24x36", StringComparison.CurrentCultureIgnoreCase)) continue;
                                foreach (ObjectId attId in acBlock.AttributeCollection)
                                {
                                    var acAtt = newTransaction.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                    if (acAtt == null) continue;

                                    if (!acAtt.Tag.Equals("PROJECT_NAME", StringComparison.CurrentCultureIgnoreCase)) continue;
                                    acAtt.UpgradeOpen();
                                    strAddress = acAtt.TextString;
                                }
                            }
                            
                        }

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("imageframe", 0);

                        Vector3d width;
                        Vector3d height;

                        // Create the new image and assign it the image definition
                        using (RasterImage acRaster = new RasterImage())
                        {
                            acRaster.ImageDefId = acImgDefId;


                            // Since the images for the seals are all different sizes, we need to resize to fit the seal area.
                            if (db.Measurement == MeasurementValue.English)
                            {
                                width = new Vector3d((acRasterDefNew.ResolutionMMPerPixel.X * acRaster.ImageWidth) / imgResize, 0, 0);
                                height = new Vector3d(0, (acRasterDefNew.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / imgResize, 0);
                            }
                            else
                            {
                                width = new Vector3d(acRasterDefNew.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0);
                                height = new Vector3d(0, acRasterDefNew.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0);
                            }

                            // Define the position for the image 
                            Point3d insPt = new Point3d(dSeal1PosX, dSeal1PosY, 0.0);

                            // Define and assign a coordinate system for the image's orientation
                            CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
                            acRaster.Orientation = coordinateSystem;

                            // Set the rotation angle for the image
                            acRaster.Rotation = 0;

                            // Add the new object to the block table record and the transaction
                            acBlkTblRec.AppendEntity(acRaster);
                            newTransaction.AddNewlyCreatedDBObject(acRaster, true);


                            // Connect the raster definition and image together so the definition
                            // does not appear as "unreferenced" in the External References palette.
                            RasterImage.EnableReactors(true);
                            acRaster.AssociateRasterDef(acRasterDefNew);

                            if (bRasterDefCreated)
                            {
                                acRasterDefNew.Dispose();
                            }
                        }

                        // Create the new image and assign it the image definition

                        using (RasterImage acRaster = new RasterImage())
                        {
                            acRaster.ImageDefId = acImgDefId2;

                            // Use ImageWidth and ImageHeight to get the size of the image in pixels (1024 x 768).
                            // Use ResolutionMMPerPixel to determine the number of millimeters in a pixel so you 
                            // can convert the size of the drawing into other units or millimeters based on the 
                            // drawing units used in the current drawing.

                            // Define the width and height of the image



                            // Check to see if the measurement is set to English (Imperial) or Metric units
                            if (db.Measurement == MeasurementValue.English)
                            {
                                width = new Vector3d((acRasterDefNew2.ResolutionMMPerPixel.X * acRaster.ImageWidth) / imgResize, 0, 0);
                                height = new Vector3d(0, (acRasterDefNew2.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / imgResize, 0);
                            }
                            else
                            {
                                width = new Vector3d(acRasterDefNew2.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0);
                                height = new Vector3d(0, acRasterDefNew2.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0);
                            }

                            // Define the position for the image 
                            Point3d insPt = new Point3d(1400.0, 1800.0, 0.0);

                            // Define and assign a coordinate system for the image's orientation
                            CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
                            acRaster.Orientation = coordinateSystem;

                            // Set the rotation angle for the image
                            acRaster.Rotation = 0;

                            // Add the new object to the block table record and the transaction
                            acBlkTblRec.AppendEntity(acRaster);
                            newTransaction.AddNewlyCreatedDBObject(acRaster, true);


                            // Connect the raster definition and image together so the definition
                            // does not appear as "unreferenced" in the External References palette.
                            RasterImage.EnableReactors(true);
                            acRaster.AssociateRasterDef(acRasterDefNew2);

                            if (bRasterDefCreated)
                            {
                                acRasterDefNew2.Dispose();
                            }
                        }

                        // Create the new image and assign it the image definition
                        //Hack to see if we have stuff on one of the detail pages.
                        if (ObjectsOnGarageLayout() > 25)
                        {
                            using (RasterImage acRaster = new RasterImage())
                            {
                                acRaster.ImageDefId = acImgDefId3;


                                // Check to see if the measurement is set to English (Imperial) or Metric units
                                if (db.Measurement == MeasurementValue.English)
                                {
                                    width = new Vector3d((acRasterDefNew3.ResolutionMMPerPixel.X * acRaster.ImageWidth) / imgResize, 0, 0);
                                    height = new Vector3d(0, (acRasterDefNew3.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / imgResize, 0);
                                }
                                else
                                {
                                    width = new Vector3d(acRasterDefNew3.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0);
                                    height = new Vector3d(0, acRasterDefNew3.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0);
                                }

                                // Define the position for the image 
                                Point3d insPt = new Point3d(-150, 750.0, 0.0);

                                // Define and assign a coordinate system for the image's orientation
                                CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
                                acRaster.Orientation = coordinateSystem;

                                // Set the rotation angle for the image
                                acRaster.Rotation = 0;

                                // Add the new object to the block table record and the transaction
                                acBlkTblRec.AppendEntity(acRaster);
                                newTransaction.AddNewlyCreatedDBObject(acRaster, true);


                                // Connect the raster definition and image together so the definition
                                // does not appear as "unreferenced" in the External References palette.
                                RasterImage.EnableReactors(true);
                                acRaster.AssociateRasterDef(acRasterDefNew3);

                                if (bRasterDefCreated)
                                {
                                    acRasterDefNew3.Dispose();
                                }
                            }
                        }

                        if (!newTextStyleTable.Has("ErikUeber"))  
                        {
                            newTextStyleTable.UpgradeOpen();
                            TextStyleTableRecord newTextStyleTableRecord = new TextStyleTableRecord();
                            newTextStyleTableRecord.FileName = "EUEBER.TTF";
                            newTextStyleTableRecord.Name = "ErikUeber";
                            newTextStyleTable.Add(newTextStyleTableRecord);
                            newTransaction.AddNewlyCreatedDBObject(newTextStyleTableRecord, true);
                        }

                        if (!newTextStyleTable.Has("MikeLynch"))  
                        {
                            newTextStyleTable.UpgradeOpen();
                            TextStyleTableRecord newTextStyleTableRecord = new TextStyleTableRecord();
                            newTextStyleTableRecord.FileName = "MIKELYNCH.TTF";
                            newTextStyleTableRecord.Name = "MikeLynch";
                            newTextStyleTable.Add(newTextStyleTableRecord);
                            newTransaction.AddNewlyCreatedDBObject(newTextStyleTableRecord, true);
                        }

                        if (!newTextStyleTable.Has("Christine"))  // !!! The TextStyle does not follow the naming convention of the other TextStyles !!!
                        {
                            newTextStyleTable.UpgradeOpen();
                            TextStyleTableRecord newTextStyleTableRecord = new TextStyleTableRecord();
                            newTextStyleTableRecord.FileName = "ccburney.TTF";
                            newTextStyleTableRecord.Name = "Christine";
                            newTextStyleTable.Add(newTextStyleTableRecord);
                            newTransaction.AddNewlyCreatedDBObject(newTextStyleTableRecord, true);
                        }

                        //set our signer
                        String strTextStyle = "ErikUeber";

                        if (radioButton5.Checked == true)
                        {
                            strTextStyle = "MikeLynch";
                        }

                        if (radioButton4.Checked == true)
                        {
                            strTextStyle = "Christine";
                        }

                        //signature
                        DBText newDBText = new DBText();
                        newDBText.SetDatabaseDefaults();
                        newDBText.Position = new Point3d(1400, 700, 0);
                        newDBText.Height = 20.0;
                        newDBText.TextString = "A";
                        newDBText.TextStyleId = newTextStyleTable[strTextStyle];
                        newBlockTableRecord.AppendEntity(newDBText);
                        newTransaction.AddNewlyCreatedDBObject(newDBText, true);

                        //signature
                        DBText newDBText2 = new DBText();
                        newDBText2.SetDatabaseDefaults();
                        newDBText2.Position = new Point3d(1400, 1750, 0);
                        newDBText2.Height = 20.0;
                        newDBText2.TextString = "A";
                        newDBText2.TextStyleId = newTextStyleTable[strTextStyle];

                        newBlockTableRecord.AppendEntity(newDBText2);
                        newTransaction.AddNewlyCreatedDBObject(newDBText2, true);

                        //if there's stuff on the Garage layout page, add a sig
                        if (ObjectsOnGarageLayout() > 25)
                        {
                            DBText newDBText3 = new DBText();
                            newDBText3.SetDatabaseDefaults();
                            newDBText3.Position = new Point3d(-150, 700, 0);
                            newDBText3.Height = 20.0;
                            newDBText3.TextString = "A";

                            newDBText3.TextStyleId = newTextStyleTable[strTextStyle];

                            newBlockTableRecord.AppendEntity(newDBText3);
                            newTransaction.AddNewlyCreatedDBObject(newDBText3, true);
                        }

                        //date
                        DBText newDBTextDate = new DBText();
                        newDBTextDate.SetDatabaseDefaults();
                        newDBTextDate.Position = new Point3d(1430, 1735, 0);
                        newDBTextDate.Height = 8.0;
                        newDBTextDate.TextString = strDate;
                        newDBTextDate.TextStyleId = newTextStyleTable[strTextStyle];

                        newBlockTableRecord.AppendEntity(newDBTextDate);
                        newTransaction.AddNewlyCreatedDBObject(newDBTextDate, true);

                        //date
                        DBText newDBTextDate2 = new DBText();
                        newDBTextDate2.SetDatabaseDefaults();
                        newDBTextDate2.Position = new Point3d(1430, 675, 0);
                        newDBTextDate2.Height = 8.0;
                        newDBTextDate2.TextString = strDate;
                        newDBTextDate2.TextStyleId = newTextStyleTable[strTextStyle];

                        newBlockTableRecord.AppendEntity(newDBTextDate2);
                        newTransaction.AddNewlyCreatedDBObject(newDBTextDate2, true);

                        //if there are items on the Garage Layout, seal it, too
                        if (ObjectsOnGarageLayout() > 25)
                        {
                            //date
                            DBText newDBTextDate3 = new DBText();
                            newDBTextDate3.SetDatabaseDefaults();
                            newDBTextDate3.Position = new Point3d(-120, 675, 0);
                            newDBTextDate3.Height = 8.0;
                            newDBTextDate3.TextString = strDate;

                            newDBTextDate3.TextStyleId = newTextStyleTable[strTextStyle];

                            newBlockTableRecord.AppendEntity(newDBTextDate3);
                            newTransaction.AddNewlyCreatedDBObject(newDBTextDate3, true);
                        }

                        newTransaction.Commit();

                        //If the deliverable is a DWF or a PDF
                        if (strFormat == "DWF" || strFormat == "PDF")
                        {
                            //Publish
                            PublishLayouts(strFormat);
                            
                            //doc.CloseAndSave(db.Filename);  <- use this to Close the file after sending
                            sendmail();
                        }

                    }
                }
                else
                {
                    acDocMgr.MdiActiveDocument.Editor.WriteMessage("File " + strFileName + " does not exist.");
                }
            }

        }


        private void addRasterImage(RasterImage ri, double x, double y)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        
        static public Int32 ObjectsOnGarageLayout()
        {
            Int32 iCount = 0;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Create a crossing window from (2,2,0) to (10,8,0)
            PromptSelectionResult acSSPrompt;
            acSSPrompt = ed.SelectCrossingWindow(new Point3d(-1500, 0, 0), new Point3d(0, 800, 0));

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;

                iCount = acSSet.Count;
            }

            return (iCount);
        }

        static public bool isLargeSize()
        {
            bool bIsLargeSize = false;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            LayoutManager layoutMgr = LayoutManager.Current;


            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = tr.GetObject(
                                    db.LayoutDictionaryId,
                                    OpenMode.ForRead,
                                    false
                                  ) as DBDictionary;

                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    ObjectId layoutId = entry.Value;

                    Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                    
                    //The larger sized drawing files have this layout, but the regular size does not 
                    if (layout.LayoutName == "S1 - 48")
                    {
                        bIsLargeSize = true;
                    }
                }
                tr.Commit();
            }

            return (bIsLargeSize);
        }

        static public Int32 Is64()
        {

            Int32 iCount = 0;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Create a crossing window from (2,2,0) to (10,8,0)
            PromptSelectionResult acSSPrompt;
            acSSPrompt = ed.SelectCrossingWindow(new Point3d(760, 1800, 0), new Point3d(2280, 2880, 0));

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;

                iCount = acSSet.Count;
            }

            return (iCount);

        }

        public static ArrayList GetLayoutIdList(Database db)
        {
            ArrayList layoutList = new ArrayList();

            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm =
                db.TransactionManager;

            using (Transaction myT = tm.StartTransaction())
            {
                DBDictionary dic =
                    (DBDictionary)tm.GetObject(db.LayoutDictionaryId,
                                        OpenMode.ForRead, false);
                DbDictionaryEnumerator index = dic.GetEnumerator();

                while (index.MoveNext())
                    layoutList.Add(index.Current.Value);

                myT.Commit();
            }

            return layoutList;
        }

    //Simple Plot is no longer used. PublishLayouts does all the work now
    static public void SimplePlot(String strJobNumber)
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;
      Database db = doc.Database;

      String strFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + strJobNumber + ".dwf";
      Transaction tr = db.TransactionManager.StartTransaction();

      using (tr)
      {
        BlockTable bt =
          (BlockTable)tr.GetObject(
            db.BlockTableId,
            OpenMode.ForRead
          );

        PlotInfo pi = new PlotInfo();
        PlotInfoValidator piv =
          new PlotInfoValidator();
        piv.MediaMatchingPolicy =
          MatchingPolicy.MatchEnabled;

        // A PlotEngine does the actual plotting
        // (can also create one for Preview)

        if (PlotFactory.ProcessPlotState ==
            ProcessPlotState.NotPlotting)
        {
                    PlotEngine pe =
                      PlotFactory.CreatePublishEngine();
                    using (pe)
                    {
                        // Collect all the paperspace layouts
                        // for plotting
                        ObjectIdCollection layoutsToPlot =
                          new ObjectIdCollection();
                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr =
                              (BlockTableRecord)tr.GetObject(
                                btrId,
                                OpenMode.ForRead
                              );
                            if (btr.IsLayout &&
                                btr.Name.ToUpper() !=
                                  BlockTableRecord.ModelSpace.ToUpper())
                            {
                                layoutsToPlot.Add(btrId);
                            }
                        }
                        // Create a Progress Dialog to provide info
                        // and allow thej user to cancel
                        PlotProgressDialog ppd =
                          new PlotProgressDialog(
                            false,
                            layoutsToPlot.Count,
                            true
                          ); 
            using (ppd)
            {

                            int numSheet = 1;
                            foreach (ObjectId btrId in layoutsToPlot)
                            {
                                
                                BlockTableRecord btr =
                                  (BlockTableRecord)tr.GetObject(
                                    btrId,
                                    OpenMode.ForRead
                                  );
                                Layout lo =
                                                  (Layout)tr.GetObject(
                                                    btr.LayoutId,
                                                    OpenMode.ForRead
                                                  );

                                // We need a PlotSettings object
                                // based on the layout settings
                                // which we then customize

                                PlotSettings ps =
                                  new PlotSettings(lo.ModelType);
                                ps.CopyFrom(lo);

                                // The PlotSettingsValidator helps
                                // create a valid PlotSettings object

                                PlotSettingsValidator psv =
                                  PlotSettingsValidator.Current;

                                // We'll plot the extents, centered and
                                // scaled to fit

                                psv.SetPlotType(
                                  ps,
                                Autodesk.AutoCAD.DatabaseServices.PlotType.Extents
                                );
                                psv.SetUseStandardScale(ps, true);
                                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                                psv.SetPlotCentered(ps, true);

                                // We'll use the standard DWFx PC3, as
                                // this supports multiple sheets

                                psv.SetPlotConfigurationName(
                                  ps,
                                  "DWFx ePlot (XPS Compatible).pc3",
                                  "ANSI_A_(8.50_x_11.00_Inches)"
                                );

                                // We need a PlotInfo object
                                // linked to the layout

                                pi.Layout = btr.LayoutId;

                                // Make the layout we're plotting current

                                LayoutManager.Current.CurrentLayout =
                                  lo.LayoutName;

                                // We need to link the PlotInfo to the
                                // PlotSettings and then validate it

                                pi.OverrideSettings = ps;
                                piv.Validate(pi);

                                if (numSheet == 1)
                                {
                                    ppd.set_PlotMsgString(
                                      PlotMessageIndex.DialogTitle,
                                      "Custom Plot Progress"
                                    );
                                    ppd.set_PlotMsgString(
                                      PlotMessageIndex.CancelJobButtonMessage,
                                      "Cancel Job"
                                    );
                                    ppd.set_PlotMsgString(
                                      PlotMessageIndex.CancelSheetButtonMessage,
                                      "Cancel Sheet"
                                    );
                                    ppd.set_PlotMsgString(
                                      PlotMessageIndex.SheetSetProgressCaption,
                                      "Sheet Set Progress"
                                    );
                                    ppd.set_PlotMsgString(
                                      PlotMessageIndex.SheetProgressCaption,
                                      "Sheet Progress"
                                    );
                                    ppd.LowerPlotProgressRange = 0;
                                    ppd.UpperPlotProgressRange = 100;
                                    ppd.PlotProgressPos = 0;

                                    // Let's start the plot, at last

                                    ppd.OnBeginPlot();
                                    ppd.IsVisible = true;
                                    pe.BeginPlot(ppd, null);

                                    // We'll be plotting a single document
                                    
                                    pe.BeginDocument(
                                      pi,
                                      doc.Name,
                                      null,
                                      1,
                                      true, // Let's plot to file
                                      strFileName
                                    );
                                    

                                }
                                // Which may contain multiple sheets

                                ppd.StatusMsgString =
                                      "Plotting " +
                                      doc.Name.Substring(
                                        doc.Name.LastIndexOf("\\") + 1
                                      ) +
                                      " - sheet " + numSheet.ToString() +
                                      " of " + layoutsToPlot.Count.ToString();

                                    ppd.OnBeginSheet();

                                    ppd.LowerSheetProgressRange = 0;
                                    ppd.UpperSheetProgressRange = 100;
                                    ppd.SheetProgressPos = 0;

                                    PlotPageInfo ppi = new PlotPageInfo();
                                    pe.BeginPage(
                                      ppi,
                                      pi,
                                      (numSheet == layoutsToPlot.Count),
                                      null
                                    );
                                    pe.BeginGenerateGraphics(null);
                                    ppd.SheetProgressPos = 50;
                                    pe.EndGenerateGraphics(null);

                                    // Finish the sheet
                                    pe.EndPage(null);
                                    ppd.SheetProgressPos = 100;
                                    ppd.OnEndSheet();
                                
                                numSheet++;
                            }

                            // Finish the document

                            pe.EndDocument(null);

                            // And finish the plot

                            ppd.PlotProgressPos = 100;
                            ppd.OnEndPlot();
                            pe.EndPlot(null);
                        }
                    }
                
            }else
        {
          MessageBox.Show("another print to DWF in progress");
        }
      }
    }

    public static void PublishLayouts(String strOutputType)
    {
        //There are a number of layouts in the dwg files, but we only want to publish 2 of the layouts for regular drawings and 3 layouts for the larger size
        using (DsdEntryCollection dsdDwgFiles = new DsdEntryCollection())
        {
            if (isLargeSize() == true)
            {
                // Define the first layout
                //this one always gets added
                using (DsdEntry dsdDwgFile1 = new DsdEntry())
                {
                    dsdDwgFile1.DwgName = strFileName;
                    dsdDwgFile1.Layout = "DETAILS";
                    dsdDwgFile1.Title = "DETAILS";
                    dsdDwgFile1.Nps = "";
                    dsdDwgFile1.NpsSourceDwg = "";

                    dsdDwgFiles.Add(dsdDwgFile1);
                }

                //2 options here. Get the right one for page 2
                if (Is64() > 100)
                {
                    using (DsdEntry dsdDwgFile2 = new DsdEntry())
                    {
                        dsdDwgFile2.DwgName = strFileName;
                        dsdDwgFile2.Layout = "S1 - 64";
                        dsdDwgFile2.Title = "S1 - 64";
                        dsdDwgFile2.Nps = "";
                        dsdDwgFile2.NpsSourceDwg = "";

                        dsdDwgFiles.Add(dsdDwgFile2);
                    }
                }
                else
                {
                    using (DsdEntry dsdDwgFile2 = new DsdEntry())
                    {
                        dsdDwgFile2.DwgName = strFileName;
                        dsdDwgFile2.Layout = "S1 - 48";
                        dsdDwgFile2.Title = "S1 - 48";
                        dsdDwgFile2.Nps = "";
                        dsdDwgFile2.NpsSourceDwg = "";

                        dsdDwgFiles.Add(dsdDwgFile2);
                    }
                }
            }
            else
            {
                // this deliverable has 3 layouts

                // always add this layout
                using (DsdEntry dsdDwgFile1 = new DsdEntry())
                {
                    dsdDwgFile1.DwgName = strFileName;
                    dsdDwgFile1.Layout = "1-Page 1";
                    dsdDwgFile1.Title = "1-Page 1";
                    dsdDwgFile1.Nps = "";
                    dsdDwgFile1.NpsSourceDwg = "";

                    dsdDwgFiles.Add(dsdDwgFile1);
                }

                // if we have stuff on the garage layout, add it
                if (ObjectsOnGarageLayout() > 50)
                {
                    using (DsdEntry dsdDwgFile2 = new DsdEntry())
                    {
                        dsdDwgFile2.DwgName = strFileName;
                        dsdDwgFile2.Layout = "2-Garage";
                        dsdDwgFile2.Title = "2-Garage";
                        dsdDwgFile2.Nps = "";
                        dsdDwgFile2.NpsSourceDwg = "";

                        dsdDwgFiles.Add(dsdDwgFile2);
                    }
                }

                // add the detail page
                using (DsdEntry dsdDwgFile3 = new DsdEntry())
                {
                    dsdDwgFile3.DwgName = strFileName;
                    dsdDwgFile3.Layout = "3-Detail Page";
                    dsdDwgFile3.Title = "3-Detail Page";
                    dsdDwgFile3.Nps = "";
                    dsdDwgFile3.NpsSourceDwg = "";

                    dsdDwgFiles.Add(dsdDwgFile3);
                }
            }

            // Set the properties for the DSD file and then write it out
            using (DsdData dsdFileData = new DsdData())
            {
  
                // Set the target location for publishing
                if (strOutputType == "PDF")
                {
                    dsdFileData.DestinationName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + strAddress + ".pdf";
                    dsdFileData.ProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
                    dsdFileData.SheetType = SheetType.MultiPdf;
                    strDeliveryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + strAddress + ".pdf";
                }
                else
                {
                    dsdFileData.DestinationName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + strAddress + ".dwf";
                    dsdFileData.ProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
                    dsdFileData.SheetType = SheetType.MultiDwf;
                    strDeliveryFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + strAddress + ".dwf";

                }

                // Set the drawings that should be added to the publication
                dsdFileData.SetDsdEntryCollection(dsdDwgFiles);

                //clean out any old batch files
                String strBatchFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\myBatch.txt";
                if (File.Exists(strBatchFile))
                {
                    File.Delete(strBatchFile);
                }

                //clean out any old dsd files
                String strDSDFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\batchdrawings2.dsd";
                if (File.Exists(strDSDFile))
                {
                    File.Delete(strDSDFile);
                }

                // Set the general publishing properties
                dsdFileData.LogFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\myBatch.txt";

                // Create the DSD file
                dsdFileData.WriteDsd(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\batchdrawings2.dsd");

                String dsdFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\batchdrawings2.dsd";

                System.IO.StreamReader sr = new System.IO.StreamReader(dsdFile);

                string str = sr.ReadToEnd();
                sr.Close();

                //this suppresses prompts while publishing runs
                str = str.Replace("PromptForDwfName=TRUE", "PromptForDwfName=FALSE");

                System.IO.StreamWriter sw = new System.IO.StreamWriter(dsdFile);

                sw.Write(str);
                sw.Close();


                try
                {
                    // Publish the specified drawing files in the DSD file,

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                    using (DsdData dsdDataFile = new DsdData())
                    {
                        dsdDataFile.ReadDsd(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\batchdrawings2.dsd");

                            // Get the DWG to PDF.pc3 and use it as a 
                            // device override for all the layouts
                            if (strOutputType == "PDF")
                            {
                                PlotConfig acPlCfg = PlotConfigManager.SetCurrentConfig("DWG to PDF.PC3");
                                Autodesk.AutoCAD.ApplicationServices.Application.Publisher.PublishExecute(dsdDataFile, acPlCfg);
                            }
                            else
                            {
                                PlotConfig acPlCfg = PlotConfigManager.SetCurrentConfig("DWF6 ePlot.pc3");
                                Autodesk.AutoCAD.ApplicationServices.Application.Publisher.PublishExecute(dsdDataFile, acPlCfg);
                            }
                    }

                    //added so that is rests for a second before it tries to email out the file.
                    Thread.Sleep(1000);

                }
                catch (Autodesk.AutoCAD.Runtime.Exception es)
                {
                    System.Windows.Forms.MessageBox.Show(es.Message);
                }
            }
        }
    }



    //public static void sendmail(object sender, Autodesk.AutoCAD.Publishing.PublishEventArgs e)
    public static void sendmail()
    {

        try
        {
            //connect to the database
            String strConn = "Server=mlawdb.cja22lachoyz.us-west-2.rds.amazonaws.com;Database=MLAW_MS;User Id=sa;Password=!sd2the2power!;";
            SqlConnection conn = new SqlConnection(strConn);
            conn.Open();

            DataSet ds = new DataSet();

            //update the order status
            SqlCommand sqlComm = new SqlCommand("Finish_Delivery", conn);
            sqlComm.Parameters.AddWithValue("@MLAW_Number", strJobNumber);

            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.ExecuteNonQuery();

            //upload the delivered file to the database
            byte[] bytes = File.ReadAllBytes(strDeliveryFile);
            SqlParameter fileP = new SqlParameter("@file", SqlDbType.VarBinary);
            fileP.Value = bytes;

            SqlParameter sqlOrderId = new SqlParameter("@MLAW_Number", SqlDbType.VarChar);
            sqlOrderId.Value = strJobNumber;

            SqlParameter sqlFileName = new SqlParameter("@File_Name", SqlDbType.VarChar);
            int iPos = strDeliveryFile.LastIndexOf("\\") + 1;
            String strFileInDB = strDeliveryFile.Substring(iPos, (strDeliveryFile.Length - iPos));

            sqlFileName.Value = strFileInDB;

            SqlCommand myCommand = new SqlCommand();
            myCommand.Parameters.Add(fileP);
            myCommand.Parameters.Add(sqlOrderId);
            myCommand.Parameters.Add(sqlFileName);
  
            myCommand.Connection = conn;
            myCommand.CommandText = "Insert_Order_File_2";
            myCommand.CommandType = CommandType.StoredProcedure;
            myCommand.ExecuteNonQuery();
            

            
            //get the email addresses of the customer contacts
            DataSet dsEmails = new DataSet();
            sqlComm = new SqlCommand("Get_Order_Client_Delivery_Emails", conn);
            sqlComm.Parameters.AddWithValue("@MLAW_Number", strJobNumber);
            sqlComm.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = sqlComm;

            da.Fill(dsEmails);

            conn.Close();
            

            //Connect to the Exchange server
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new WebCredentials("Distribute@mlaw-eng.com", "D!str!but3");  // this does not work because the user is authenticated already with the login they used to logon to their PC
            service.UseDefaultCredentials = true;  // so, we have to use their default credentials (whichever user they used to logon to the PC)
            service.AutodiscoverUrl("Distribute@mlaw-eng.com", RedirectionUrlValidationCallback);

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            //Create our email message
            EmailMessage email = new EmailMessage(service);
           

            email.BccRecipients.Add("joe@kickasstek.com");
            email.BccRecipients.Add("CDYork@mlaw-eng.com");

            //Add recipients
            foreach (DataRow dr in dsEmails.Tables[0].Rows)
            {
                email.ToRecipients.Add(dr["Email_Address"].ToString().Trim());
            }

            
            DateTime dtNow = DateTime.Now;

            email.Subject = dtNow.ToString("MM-dd-yyyy") + " - " + strAddress;
            email.Body = new MessageBody("Please find it attached");
            email.Attachments.AddFileAttachment(strDeliveryFile);

            //send the email
            email.Send();
           

        }
        catch (System.Exception ex)
        {
            // The MLAW network randomly closes connections from time to time. If that happens, retry
            if (ex.Message.IndexOf("connection") > -1)
            {
                Thread.Sleep(3000);
                sendmail();
            }
            else
            {
                MessageBox.Show(ex.Message);
            }
        }
        
    }

    private static bool RedirectionUrlValidationCallback(string redirectionUrl)
    {
        // The default for the validation callback is to reject the URL.
        bool result = false;

        Uri redirectionUri = new Uri(redirectionUrl);

        if (redirectionUri.Scheme == "https")
        {
            result = true;
        }
        return result;
    }

    private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {

    }

    private void button2_Click(object sender, EventArgs e)
    {
        deliverJobs();
    }

    private void button3_Click(object sender, EventArgs e)
    {
        ListBox.SelectedObjectCollection selectedItems = new ListBox.SelectedObjectCollection(lstSelected);
        selectedItems = lstSelected.SelectedItems;

        if (lstSelected.SelectedIndex != -1)
        {
            for (int i = selectedItems.Count - 1; i >= 0; i--)
                lstSelected.Items.Remove(selectedItems[i]);
        }
    }
  
    }


}
