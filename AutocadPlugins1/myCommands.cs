// (C) Copyright 2014 by Microsoft 
//
using System;
using System.Linq;
using System.IO;
using System.Reflection;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;

using acap = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Collections.Generic;
using Autodesk.AutoCAD.Colors;
using System.Data.OleDb;


// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(AutocadPlugins1.MyCommands))]

namespace AutocadPlugins1
{

    public class MyCommands
    {
        public double dMaxX = 0;
        public double dMinX = 0;
        public double dMaxY = 0;
        public double dMinY = 0;


        public ObjectId[] selectedIds;
        public ArrayList verticalRects;
        public ArrayList horizontalRects;

        public ArrayList horizBoundLines;
        public ArrayList vertBoundLines;

        public ObjectIdCollection allHorizLines;
        public ObjectIdCollection allVertLines;


        [CommandMethod("DELIVERY")]
        public void doDelivery()
        {
            Delivery frmDelivery = new Delivery();
            frmDelivery.Width = 1300;
            frmDelivery.Height = 600;
            Application.ShowModelessDialog(null,frmDelivery, false);

        }
        
        [CommandMethod("TTL")]
        public void getTitleInfo()
        {
            //Autocad Parameters
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            String strFileName = db.Filename;

            //Get what we need for a job number
            Int32 iPos1 = strFileName.LastIndexOf("\\") + 1;
            Int32 iPos2 = strFileName.LastIndexOf(".");
            String strJobNumber = strFileName.Substring(iPos1, (iPos2 - iPos1));

            //Create the path to the file in use
            string drive = Path.GetPathRoot("M:");   
            String strPath = "";

            String strCurYear = "20" + strJobNumber.Substring(0, 2);

            if (!Directory.Exists(drive))
            {
                //Dev test path
                strPath = "C:\\Users\\Joe\\Downloads\\";
            }
            else
            {
                //Live MLAW environment
                strPath = "M:\\SFR Foundations\\" + strCurYear + " Final Designs\\";
            }

            //Database Connection string
            String strConn = "Server=mlawdb.cja22lachoyz.us-west-2.rds.amazonaws.com;Database=MLAW_MS;User Id=sa;Password=!sd2the2power!;";

            //If Autocad has autosaved the file it adds junk to the end of the filename. Remove it.
            Int32 iPos3 = strJobNumber.IndexOf("_");
            if (iPos3 != -1)
            {
                strJobNumber = strJobNumber.Substring(0, iPos3);
                strJobNumber = strJobNumber.Replace("_", "");
            }

            //output to the user
            ed.WriteMessage("Titleing for job number: " + strJobNumber + "\n");

            //Get our order information from the database
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                SqlCommand sqlComm = new SqlCommand("Get_Order_By_MLAW_Number", conn);
                sqlComm.Parameters.AddWithValue("@MLAW_Number", strJobNumber);

                sqlComm.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlComm;

                da.Fill(ds);
            }

            //If we have results, add them to the title block
            if (ds.Tables[0].Rows.Count > 0)
            {
                DataRow dr = ds.Tables[0].Rows[0];

                using (Transaction tTrAct = doc.TransactionManager.StartTransaction())
                {
                    doc.TransactionManager.EnableGraphicsFlush(true);
                    BlockTable tBlTab = tTrAct.GetObject(doc.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    IEnumerator tBlTabEnum = tBlTab.GetEnumerator();
                    while (tBlTabEnum.MoveNext())
                    {
                        BlockTableRecord tBlTabRec = tTrAct.GetObject((ObjectId)tBlTabEnum.Current, OpenMode.ForWrite) as BlockTableRecord;
                        IEnumerator tBlTabRecEnum = tBlTabRec.GetEnumerator();
                        while (tBlTabRecEnum.MoveNext())
                        {

                            ObjectId tObjID = (ObjectId)tBlTabRecEnum.Current;
                            if (tBlTabRec.Name == "TITLE_BLOCK")
                            {
                                if ((tObjID.IsValid == true) && (tObjID.IsErased == false) && (tObjID.ObjectClass.DxfName.ToUpper() == "ATTDEF"))
                                {
                                    String strDate = DateTime.Now.ToString("d");
                                    String[] arrUser = Application.GetSystemVariable("LOGINNAME").ToString().Split(' ');
                                    String strUser = "";

                                    if (arrUser.Length > 1)
                                    {
                                        strUser = arrUser[0].Substring(0, 1) + arrUser[1].Substring(0, 1);
                                    }
                                    else
                                    {
                                        strUser = arrUser[0].Substring(0, 1);
                                    }

                                    AttributeDefinition tAttDef = tTrAct.GetObject(tObjID, OpenMode.ForRead) as AttributeDefinition;
                                    EditBlockAtt("TITLE_BLOCK", "JOB_NUMBER", dr["MLAW_Number"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "CLIENT", dr["Client_Full_Name"].ToString().ToUpper());
                                    EditBlockAtt("TITLE_BLOCK", "SUBDIVISION", dr["Subdivision_Name"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "LOT", dr["Lot"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "BLOCK", dr["Block"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "PHASE", dr["Phase"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "SECTION", dr["Section"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "ADDRESS", dr["Address"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "CITY", dr["City"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "DATE", strDate);
                                    EditBlockAtt("TITLE_BLOCK", "DRAWN_BY", strUser );
                                    EditBlockAtt("TITLE_BLOCK", "PLAN_NO.", dr["Plan_Number"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "SOILS_DATA_SOURCE", dr["Soils_Data_Source"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "SOILS_DATA_DATE", dr["Visual_Geotec_Date"].ToString());
                                    EditBlockAtt("TITLE_BLOCK", "PI", dr["PI"].ToString());

                                    Double dEMCTR = Convert.ToDouble(dr["Em_ctr"]);
                                    Double dEMEDG = Convert.ToDouble(dr["Em_edg"]);
                                    Double dYMCTR = Convert.ToDouble(dr["Ym_ctr"]);
                                    Double dYMEDG = Convert.ToDouble(dr["Ym_edg"]);


                                    EditBlockAtt("TITLE_BLOCK", "EMCTR", dEMCTR.ToString("n2"));
                                    EditBlockAtt("TITLE_BLOCK", "EMEDG", dEMEDG.ToString("n2"));
                                    EditBlockAtt("TITLE_BLOCK", "YMCTR", dYMCTR.ToString("n2"));
                                    EditBlockAtt("TITLE_BLOCK", "YMEDG", dYMEDG.ToString("n2"));

                                }
                            }
                        }
                    }

                    tTrAct.TransactionManager.QueueForGraphicsFlush();
                    doc.TransactionManager.FlushGraphics();

                    tTrAct.Commit();
                    ed.Regen();

                }

            }
            else
            {
                //if there are no results, tell the user
                ed.WriteMessage("UNABLE TO FIND A JOB THAT MATCHED THIS JOB NUMBER: " + strJobNumber + "\n");
            }
            
        }
        
    //function to edit a title block
    public void EditBlockAtt(String blockName, String attName, String newValue)
    {

        var acDb = HostApplicationServices.WorkingDatabase;
        var acEd = acap.DocumentManager.MdiActiveDocument.Editor;

        using (var acTrans = acDb.TransactionManager.StartTransaction())
        {
            var acBlockTable = acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (acBlockTable == null) return;

            var acBlockTableRecord =
                acTrans.GetObject(acBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
            if (acBlockTableRecord == null) return;

            foreach (var blkId in acBlockTableRecord)
            {
                var acBlock = acTrans.GetObject(blkId, OpenMode.ForRead) as BlockReference;
                if (acBlock == null) continue;
                if (!acBlock.Name.Equals(blockName, StringComparison.CurrentCultureIgnoreCase)) continue;
                foreach (ObjectId attId in acBlock.AttributeCollection)
                {
                    var acAtt = acTrans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (acAtt == null) continue;

                    if (!acAtt.Tag.Equals(attName, StringComparison.CurrentCultureIgnoreCase)) continue;

                    acAtt.UpgradeOpen();
                    acAtt.TextString = newValue;
                }
            }

            acTrans.Commit();
        }
    }



        
        private static ObjectIdCollection
          GetEntitiesOnLayer(string layerName)
        {
            Document doc =
              Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Build a filter list so that only entities
            // on the specified layer are selected

            TypedValue[] tvs =
              new TypedValue[1] {
            new TypedValue(
              (int)DxfCode.LayerName,
              layerName
            )
          };
            SelectionFilter sf =
              new SelectionFilter(tvs);
            PromptSelectionResult psr =
              ed.SelectAll(sf);

            if (psr.Status == PromptStatus.OK)
                return
                  new ObjectIdCollection(
                    psr.Value.GetObjectIds()
                  );
            else
                return new ObjectIdCollection();
        }
        
    }
}


