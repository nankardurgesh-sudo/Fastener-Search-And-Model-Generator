using NXOpen;
using NXOpen.Features;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NXOpen.CAE.QualityAudit.Action;

namespace SmartFastener
{
    internal class HexagonalRegularNut
    {
        public static void Create(string mSize)
        {
            Dictionary<string, string> holeDiameters = new Dictionary<string, string>
            {
                {"M1.6", "1.25"},
                {"M2", "1.6"},
                {"M2.5", "2.05"},
                {"M3", "2.5"},
                {"M4", "3.3"},
                {"M5", "4.2"},
                {"M6", "5"},
                {"M8", "6.75"},
                {"M10", "8.5"},
                {"M12", "10.25"},
                {"M16", "14"},
                {"M20", "17.5"},
                {"M24", "21"},
                {"M30", "26.5"},
                {"M36", "32"},
                {"M42", "37.5"},
                {"M48", "43"},
                {"M56", "50.5"},
                {"M64", "58"}
            };
            Dictionary<string, string> nutHeights = new Dictionary<string, string>
            {
                {"M1.6", "1.175"},
                {"M2",   "1.475"},
                {"M2.5", "1.875"},
                {"M3",   "2.275"},
                {"M4",   "3.05"},
                {"M5",   "4.55"},
                {"M6",   "5.05"},
                {"M8",   "6.62"},
                {"M10",  "8.22"},
                {"M12",  "10.585"},
                {"M16",  "14.45"},
                {"M20",  "17.45"},
                {"M24",  "20.85"},
                {"M30",  "24.95"},
                {"M36",  "30.20"},
                {"M42",  "33.20"},
                {"M48",  "37.20"},
                {"M56",  "44.20"},
                {"M64",  "50.05"}
            };
            Dictionary<string, string> widthAcrossFlat = new Dictionary<string, string>
            {
                {"M1.6", "3.11"},
                {"M2",   "3.91"},
                {"M2.5", "4.91"},
                {"M3",   "5.41"},
                {"M4",   "6.89"},
                {"M5",   "7.89"},
                {"M6",   "9.89"},
                {"M8",   "12.865"},
                {"M10",  "15.865"},
                {"M12",  "17.865"},
                {"M16",  "23.835"},
                {"M20",  "29.58"},
                {"M24",  "35.5"},
                {"M30",  "45.5"},
                {"M36",  "54.4"},
                {"M42",  "64.05"},
                {"M48",  "74.05"},
                {"M56",  "83.9"},
                {"M64",  "93.9"}
            };

            string holeDia = holeDiameters[mSize];
            string nutHeight = nutHeights[mSize];
            string WAF = widthAcrossFlat[mSize];

            Session theSession = Session.GetSession();

            try
            {
                BasePart workPart = theSession.Parts.OpenBaseDisplay(@"C:\Users\PC\Desktop\SmartFastener\HexRegularNut\HRN_Master.prt", out PartLoadStatus pls);

                pls.Dispose();

                theSession.Parts.SetWork(workPart);

                workPart.Expressions.FindObject("holeDia").RightHandSide = holeDia;
                workPart.Expressions.FindObject("nutHeight").RightHandSide = nutHeight;
                workPart.Expressions.FindObject("WAF").RightHandSide = WAF;

                Session.UndoMarkId markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Update Nut Geometry");

                theSession.UpdateManager.DoUpdate(markId);

                Part wP = theSession.Parts.Work;

                NXOpen.Features.RemoveParametersBuilder removeParametersBuilder1;
                removeParametersBuilder1 = wP.Features.CreateRemoveParametersBuilder();

                NXOpen.Body body = wP.Bodies.ToArray()[0];

                bool added1;
                added1 = removeParametersBuilder1.Objects.Add(body);

                NXObject obj1 = removeParametersBuilder1.Commit();
                removeParametersBuilder1.Destroy();

                PartSaveStatus pss = workPart.SaveAs("C:\\Users\\PC\\Desktop\\SmartFastener\\HexRegularNut\\HRN" + Excel.GetNewHRNPN() + ".prt");
                pss.Dispose();

                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

                //theSession.ListingWindow.Open();
                //theSession.ListingWindow.WriteLine("Part Generated Successfully");
            }
            catch(Exception ex)
            {
                UI.GetUI().NXMessageBox.Show("Error in Hex Nut Creation", NXMessageBox.DialogType.Error, ex.Message);
            }

        }
    }
}
