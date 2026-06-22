from pyrevit import revit, DB, forms

doc = revit.doc
# Collect all pipes in the current document
collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()

seen_geometry = set()
duplicates = []

for pipe in collector:
    loc = pipe.Location
    if not isinstance(loc, DB.LocationCurve):
        continue
        
    p1 = loc.Curve.GetEndPoint(0)
    p2 = loc.Curve.GetEndPoint(1)
    
    # Round coordinates to 3 decimals to account for micro-inaccuracies
    pt1 = (round(p1.X, 3), round(p1.Y, 3), round(p1.Z, 3))
    pt2 = (round(p2.X, 3), round(p2.Y, 3), round(p2.Z, 3))
    
    # Sort the two points so draw direction (A to B vs B to A) doesn't matter
    geom_key = tuple(sorted([pt1, pt2]))
    
    if geom_key in seen_geometry:
        duplicates.append(pipe)
    else:
        seen_geometry.add(geom_key)

# UI Output
if not duplicates:
    forms.alert("No duplicate pipes found.", title="Result")
else:
    # Format the list for the pyRevit window
    class PipeOption(forms.TemplateListItem):
        @property
        def name(self):
            return "Pipe ID: {}".format(self.item.Id)

    options = [PipeOption(p) for p in duplicates]
    
    # Display clickable window
    selected_pipes = forms.SelectFromList.show(
        options,
        title="{} Duplicate Pipes Found".format(len(duplicates)),
        multiselect=True,
        button_name="Select in Revit"
    )
    
    # Select the chosen duplicates in the Revit interface for easy deletion
    if selected_pipes:
        revit.get_selection().set_to([p.Id for p in selected_pipes])