from pyrevit import revit, DB

doc = revit.doc
selection = revit.uidoc.Selection.GetElementIds()

if not selection:
    print("Please select elements before running the script.")
else:
    with revit.Transaction("Map Level to IfcSpatialContainer"):
        updated_count = 0
        
        for el_id in selection:
            element = doc.GetElement(el_id)
            
            # Skip if the selection accidentally includes a Type instead of an Instance
            if isinstance(element, DB.ElementType):
                continue
            
            level_name = None
            
            # 1. Check if the element is hosted (e.g., Railing on a Stair)
            if hasattr(element, "HostId") and element.HostId != DB.ElementId.InvalidElementId:
                host = doc.GetElement(element.HostId)
                if host:
                    # Look for the Base Level parameter on the host
                    base_level_param = host.get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
                    if base_level_param and base_level_param.HasValue:
                        level = doc.GetElement(base_level_param.AsElementId())
                        if level:
                            level_name = level.Name
                    
                    # Fallback to the host's direct LevelId if Base Level parameter doesn't exist
                    elif hasattr(host, "LevelId") and host.LevelId != DB.ElementId.InvalidElementId:
                        level = doc.GetElement(host.LevelId)
                        if level:
                            level_name = level.Name

            # 2. Fallback to the element's own LevelId (for non-hosted elements)
            if not level_name and hasattr(element, "LevelId") and element.LevelId != DB.ElementId.InvalidElementId:
                level = doc.GetElement(element.LevelId)
                if level:
                    level_name = level.Name
            
            # 3. Write the resolved level name to IfcSpatialContainer
            if level_name:
                param = element.LookupParameter("IfcSpatialContainer")
                if param and not param.IsReadOnly:
                    param.Set(level_name)
                    updated_count += 1
                        
        print("Successfully updated IfcSpatialContainer for {} elements.".format(updated_count))