from pyrevit import revit, DB

doc = revit.doc

# Start transaction to modify elements
with revit.Transaction("Map All Stair Levels to Components"):
    updated_count = 0
    
    # 1. Collect every Stair instance in the entire model
    all_stairs = DB.FilteredElementCollector(doc).OfClass(DB.Architecture.Stairs).WhereElementIsNotElementType().ToElements()
    
    if not all_stairs:
        print("No stairs found in the model.")
    else:
        for stair in all_stairs:
            # 2. Get the Base Level name from the Stair
            level_name = None
            base_level_param = stair.get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
            
            if base_level_param and base_level_param.HasValue:
                level = doc.GetElement(base_level_param.AsElementId())
                if level:
                    level_name = level.Name
                    
            # If we successfully found the stair's level, proceed to its components
            if level_name:
                # 3. Gather all Runs, Landings, and Supports belonging to this specific stair
                component_ids = list(stair.GetStairsRuns())
                component_ids.extend(list(stair.GetStairsLandings()))
                component_ids.extend(list(stair.GetStairsSupports()))
                
                # 4. Loop through all gathered components and assign the parameter
                for comp_id in component_ids:
                    comp = doc.GetElement(comp_id)
                    if comp:
                        param = comp.LookupParameter("IfcSpatialContainer")
                        if param and not param.IsReadOnly:
                            param.Set(level_name)
                            updated_count += 1
                            
        print("Successfully updated IfcSpatialContainer for {} Stair Runs, Landings, and Supports.".format(updated_count))