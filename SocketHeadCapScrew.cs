using NXOpen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartFastener
{
    internal class SocketHeadCapScrew
    {
        public static void Create(string mSize, string length)
        {
            Dictionary<string, string> boltDiameters = new Dictionary<string, string>()
            {
                { "M1.6", "1.6" },
                { "M2",   "2.0" },
                { "M2.5", "2.5" },
                { "M3",   "3.0" },
                { "M4",   "4.0" },
                { "M5",   "5.0" },
                { "M6",   "6.0" },
                { "M8",   "8.0" },
                { "M10", "10.0" },
                { "M12", "12.0" }
            };
            Dictionary<string, string> headDiameters = new Dictionary<string, string>()
            {
                { "M1.6", "3.000" },
                { "M2",   "3.800" },
                { "M2.5", "4.500" },
                { "M3",   "5.500" },
                { "M4",   "7.000" },
                { "M5",   "8.500" },
                { "M6",   "10.000" },
                { "M8",   "13.000" },
                { "M10", "16.000" },
                { "M12", "18.000" }
            };
            Dictionary<string, string> headHeights = new Dictionary<string, string>()
            {
                { "M1.6", "1.530" },
                { "M2",   "1.930" },
                { "M2.5", "2.430" },
                { "M3",   "2.930" },
                { "M4",   "3.910" },
                { "M5",   "4.910" },
                { "M6",   "5.850" },
                { "M8",   "7.820" },
                { "M10", "9.820" },
                { "M12", "11.785" }
            };
            Dictionary<string, string> slotDepths = new Dictionary<string, string>()
            {
                { "M1.6", "0.70" },
                { "M2",   "1.00" },
                { "M2.5", "1.10" },
                { "M3",   "1.30" },
                { "M4",   "2.00" },
                { "M5",   "2.50" },
                { "M6",   "3.00" },
                { "M8",   "4.00" },
                { "M10", "5.00" },
                { "M12", "6.00" }
            };
            Dictionary<string, string> slotDiameters = new Dictionary<string, string>()
            {
                { "M1.6", "1.733" },
                { "M2",   "1.733" },
                { "M2.5", "2.303" },
                { "M3",   "2.873" },
                { "M4",   "3.443" },
                { "M5",   "4.583" },
                { "M6",   "5.723" },
                { "M8",   "6.863" },
                { "M10", "9.149" },
                { "M12", "11.429" }
            };

            string BoltDia = boltDiameters[mSize];
            string HeadDia = headDiameters[mSize];
            string HeadHeight = headHeights[mSize];
            string SlotDepth = slotDepths[mSize];
            string SlotDia = slotDiameters[mSize];

            Session theSession = Session.GetSession();

            try
            {
                BasePart workPart = theSession.Parts.OpenBaseDisplay(@"C:\Users\PC\Desktop\SmartFastener\SocketHeadCapScrew\SHCS_Master.prt", out PartLoadStatus pls);

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

                PartSaveStatus pss = workPart.SaveAs("C:\\Users\\PC\\Desktop\\SmartFastener\\SocketHeadCapScrew\\SHCS" + Excel.GetNewSHCSPN() + ".prt");
                pss.Dispose();

                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

                //theSession.ListingWindow.Open();
                //theSession.ListingWindow.WriteLine("Part Generated Successfully");
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show("Error in Socket Head Cap Screw Creation", NXMessageBox.DialogType.Error, ex.Message);
            }
        }
    }
}