using NXOpen;
using NXOpen.Features;
using NXOpenUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NXOpen.Tooling.ScrapDesignBuilder;

namespace SmartFastener
{
    internal class PlainWasher
    {
        public static void Create(string id, string od, string thickness)
        {
            Session theSession = Session.GetSession();
            try
            {
                BasePart workPart = theSession.Parts.OpenBaseDisplay("C:\\Users\\PC\\Desktop\\SmartFastener\\PlainWasher\\PW_Master.prt", out PartLoadStatus pls);

                pls.Dispose();

                theSession.Parts.SetWork(workPart);

                workPart.Expressions.FindObject("ID").RightHandSide = id;
                workPart.Expressions.FindObject("OD").RightHandSide = od;
                workPart.Expressions.FindObject("Thickness").RightHandSide = thickness;

                Session.UndoMarkId markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Update Washer Geometry");

                theSession.UpdateManager.DoUpdate(markId);

                Part wP = theSession.Parts.Work;
                
                NXOpen.Features.RemoveParametersBuilder removeParametersBuilder1;
                removeParametersBuilder1 = wP.Features.CreateRemoveParametersBuilder();

                NXOpen.Body body = wP.Bodies.ToArray()[0];

                bool added1;
                added1 = removeParametersBuilder1.Objects.Add(body);

                NXObject obj1 = removeParametersBuilder1.Commit();
                removeParametersBuilder1.Destroy();

                PartSaveStatus pss = workPart.SaveAs("C:\\Users\\PC\\Desktop\\SmartFastener\\PlainWasher\\PW" + Excel.GetNewPWPN() + ".prt");
                pss.Dispose();

                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);
                 
                //theSession.ListingWindow.Open();
                //theSession.ListingWindow.WriteLine("Part Generated Successfully");
            }

            catch(Exception ex)
            {
                UI.GetUI().NXMessageBox.Show("Error in Plain Washer Creation", NXMessageBox.DialogType.Error, ex.Message);
            }
        }
    }
}
