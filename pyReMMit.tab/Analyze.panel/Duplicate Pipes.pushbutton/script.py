from pyrevit import revit, DB, forms
from collections import defaultdict

doc = revit.doc
collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType()

seen_geometry = defaultdict(list)

for pipe in collector:
    loc = pipe.Location
    if not isinstance(loc, DB.LocationCurve):
        continue
        
    p1 = loc.Curve.GetEndPoint(0)
    p2 = loc.Curve.GetEndPoint(1)
    
    pt1 = (round(p1.X, 3), round(p1.Y, 3), round(p1.Z, 3))
    pt2 = (round(p2.X, 3), round(p2.Y, 3), round(p2.Z, 3))
    
    geom_key = tuple(sorted([pt1, pt2]))
    seen_geometry[geom_key].append(pipe)

first_pipes = []   # Originals (Lower IDs)
second_pipes = []  # Duplicates (Higher IDs)
display_items = []
pair_num = 1

for geom_key, pipes in seen_geometry.items():
    if len(pipes) > 1:
        pipes.sort(key=lambda p: p.Id.Value)
        
        first_pipes.append(pipes[0])
        label_first = "Pair {} | Pipe ID: {} (Original)".format(pair_num, pipes[0].Id)
        display_items.append({"pipe": pipes[0], "label": label_first})
        
        for pipe in pipes[1:]:
            second_pipes.append(pipe)
            label_sub = "Pair {} | Pipe ID: {} (Duplicate)".format(pair_num, pipe.Id)
            display_items.append({"pipe": pipe, "label": label_sub})
            
        pair_num += 1

if not display_items:
    forms.alert("No duplicate pipes found.", title="Result")
else:
    choices = [
        "Open list to choose manually",
        "Select ALL Duplicates (Second pipes / Higher IDs)",
        "Select ALL Originals (First pipes / Lower IDs)"
    ]
    
    user_choice = forms.CommandSwitchWindow.show(
        choices,
        message="Found {} duplicate pairs. Choose an action:".format(pair_num - 1)
    )
    
    if user_choice == "Select ALL Duplicates (Second pipes / Higher IDs)":
        revit.get_selection().set_to([p.Id for p in second_pipes])
        
    elif user_choice == "Select ALL Originals (First pipes / Lower IDs)":
        revit.get_selection().set_to([p.Id for p in first_pipes])
        
    elif user_choice == "Open list to choose manually":
        class PipeOption(forms.TemplateListItem):
            @property
            def name(self):
                return self.item["label"]

        options = [PipeOption(item) for item in display_items]

        selected_items = forms.SelectFromList.show(
            options,
            title="Duplicate Pipes Found",
            multiselect=True,
            button_name="Select in Revit"
        )
        
        if selected_items:
            revit.get_selection().set_to([item["pipe"].Id for item in selected_items])