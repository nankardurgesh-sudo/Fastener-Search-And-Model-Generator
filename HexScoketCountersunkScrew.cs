using Microsoft.Office.Interop.Excel;
using NXOpen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NXOpen.Features.VehicleDesign.PedestrianProtectionBuilder;

namespace SmartFastener
{
    internal class HexScoketCountersunkScrew
    {

        public static void Create(string mSize, string length)
        {
            Dictionary<string, string> boltDiameters = new Dictionary<string, string>()
            {
                { "M3",  "3.0" },
                { "M4",  "4.0" },
                { "M5",  "5.0" },
                { "M6",  "6.0" },
                { "M8",  "8.0" },
                { "M10", "10.0" },
                { "M12", "12.0" },
                { "M14", "14.0" },
                { "M16", "16.0" },
                { "M20", "20.0" }
            };
            Dictionary<string, string> headDiameters = new Dictionary<string, string>()
            {
                { "M3",  "6.130" },
                { "M4",  "8.245" },
                { "M5", "10.315" },
                { "M6", "12.390" },
                { "M8", "16.580" },
                { "M10", "20.810" },
                { "M12", "25.000" },
                { "M14", "28.660" },
                { "M16", "31.305" },
                { "M20", "38.185" }
            };
            Dictionary<string, string> headHeights = new Dictionary<string, string>()
            {
                { "M3",  "1.86" },
                { "M4",  "2.48" },
                { "M5",  "3.10" },
                { "M6",  "3.72" },
                { "M8",  "4.96" },
                { "M10", "6.20" },
                { "M12", "7.44" },
                { "M14", "8.40" },
                { "M16", "8.80" },
                { "M20", "10.16" }
            };
            Dictionary<string, string> slotDepths = new Dictionary<string, string>()
            {
                { "M3",  "1.10" },
                { "M4",  "1.50" },
                { "M5",  "1.90" },
                { "M6",  "2.20" },
                { "M8",  "3.00" },
                { "M10", "3.60" },
                { "M12", "4.30" },
                { "M14", "4.50" },
                { "M16", "4.80" },
                { "M20", "5.60" }
            };
            Dictionary<string, string> slotDiameters = new Dictionary<string, string>()
            {
                { "M3",  "2.303" },
                { "M4",  "2.873" },
                { "M5",  "3.443" },
                { "M6",  "4.583" },
                { "M8",  "5.723" },
                { "M10", "6.863" },
                { "M12", "9.149" },
                { "M14", "11.429" },
                { "M16", "11.429" },
                { "M20", "13.716" }
            };

            string BoltDia = boltDiameters[mSize];
            string HeadDia = headDiameters[mSize];
            string HeadHeight = headHeights[mSize];
            string SlotDepth = slotDepths[mSize];
            string SlotDia = slotDiameters[mSize];

            Session theSession = Session.GetSession();

            try
            {
                BasePart workPart = theSession.Parts.OpenBaseDisplay(@"C:\Users\PC\Desktop\SmartFastener\HexSocketCountersunkScrew\HSCS_Master.prt", out PartLoadStatus pls);

                pls.Dispose();

                theSession.Parts.SetWork(workPart);

                workPart.Expressions.FindObject("BoltDia").RightHandSide = BoltDia;
                workPart.Expressions.FindObject("HeadDia").RightHandSide = HeadDia;
                workPart.Expressions.FindObject("HeadHeight").RightHandSide = HeadHeight;
                workPart.Expressions.FindObject("SlotDepth").RightHandSide = SlotDepth;
                workPart.Expressions.FindObject("SlotDia").RightHandSide = SlotDia;
                workPart.Expressions.FindObject("Length").RightHandSide = length;

                Session.UndoMarkId markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Update Screw Geometry");

                theSession.UpdateManager.DoUpdate(markId);

                Part wP = theSession.Parts.Work;

                NXOpen.Features.RemoveParametersBuilder removeParametersBuilder1;
                removeParametersBuilder1 = wP.Features.CreateRemoveParametersBuilder();

                NXOpen.Body body = wP.Bodies.ToArray()[0];

                bool added1;
                added1 = removeParametersBuilder1.Objects.Add(body);

                NXObject obj1 = removeParametersBuilder1.Commit();
                removeParametersBuilder1.Destroy();

                PartSaveStatus pss = workPart.SaveAs("C:\\Users\\PC\\Desktop\\SmartFastener\\HexSocketCountersunkScrew\\HSCS" + Excel.GetNewHSCSPN() + ".prt");
                pss.Dispose();

                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

                //theSession.ListingWindow.Open();
                //theSession.ListingWindow.WriteLine("Part Generated Successfully");
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show("Error in Hex Scoket Countersunk Screw Creation", NXMessageBox.DialogType.Error, ex.Message);
            }
        }
    }
}
