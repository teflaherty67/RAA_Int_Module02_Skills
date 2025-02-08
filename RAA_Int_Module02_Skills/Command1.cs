using RAA_Int_Module02_Skills.Common;

namespace RAA_Int_Module02_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // 01a. Filtered Element Collector by view

            // set the current view as the active view
            View curView = curDoc.ActiveView;

            // create the collector for the active view
            FilteredElementCollector viewCol = new FilteredElementCollector(curDoc, curView.Id);

            // 01b. ElementMultiCategoryFilter

            // create the list of categories for the filter
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Doors);
            catList.Add(BuiltInCategory.OST_Rooms);
            catList.Add(BuiltInCategory.OST_Walls);
            catList.Add(BuiltInCategory.OST_Areas);

            // create the filter
            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);

            // aplly filter to collector
            viewCol.WherePasses(catFilter)
               .WhereElementIsNotElementType();

            // 1c. use LINQ to get the family symbol by name
            FamilySymbol tagDoor = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Door Tag"))
                .First();

            FamilySymbol tagRoom = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Room Tag"))
                .First();

            FamilySymbol tagWall = new FilteredElementCollector(curDoc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Wall Tag"))
                .First();

            // 2. create dictionary for tags
            Dictionary<string, FamilySymbol> d_Tags = new Dictionary<string, FamilySymbol>();

            // add keys & values to dictionary
            d_Tags.Add("Doors", tagDoor);
            d_Tags.Add("Rooms", tagRoom);
            d_Tags.Add("Walls", tagWall);

            // code snippet to retrieve particual symbol from the dictionary
            FamilySymbol curDrTag = d_Tags["Doors"];

            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Insert tags");

                foreach (Element cureElem in viewCol)
                {
                    // 3. get point from location
                    XYZ insPoint;
                    LocationCurve locCurve;
                    LocationPoint locPoint;

                    Location curLoc = cureElem.Location;

                    if (curLoc == null)
                        continue;

                    locPoint = curLoc as LocationPoint;
                    if (locPoint != null)
                    {
                        // is a location point
                        insPoint = locPoint.Point;
                    }
                    else
                    {
                        // is a locatin curve
                        locCurve = curLoc as LocationCurve;
                        Curve curCurve = locCurve.Curve;

                        insPoint = Utils.GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                    }

                    FamilySymbol curTagType = d_Tags[cureElem.Category.Name];

                    // 4. create reference to element
                    Reference curRef = new Reference(cureElem);

                    // 5a. place tag
                    IndependentTag newTag = IndependentTag.Create(curDoc, curTagType.Id, curView.Id,
                        curRef, false, TagOrientation.Horizontal, insPoint);

                    // 5b. place area tag
                    if (cureElem.Category.Name == "Areas")
                    {
                        ViewPlan curAreaPlan = curView as ViewPlan;
                        Area curArea = cureElem as Area;

                        AreaTag curAreaTag = curDoc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
                        curAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                        curAreaTag.HasLeader = false;
                    }
                }

                t.Commit();
            }


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }

}
