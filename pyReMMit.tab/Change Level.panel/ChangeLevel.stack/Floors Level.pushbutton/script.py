"""Changes host level of selected floors while maintaining their absolute position."""
__title__ = 'Change Floor Level'
__author__ = 'MM'

import clr
from Autodesk.Revit.DB import *
from pyrevit import forms
from collections import Counter

# Get current document, application and selection
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_levels():
    """Get all levels in the current document sorted from highest to lowest."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    return sorted(levels, key=lambda x: x.Elevation, reverse=True)

def get_selected_floors():
    """Get selected floor elements from current selection."""
    selection = uidoc.Selection.GetElementIds()
    return [doc.GetElement(elid) for elid in selection if isinstance(doc.GetElement(elid), Floor)]

def get_floor_info(floor):
    """Get level, offset, and absolute elevation of a floor."""
    level_id = floor.LevelId
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
        
    level = doc.GetElement(level_id)
    offset_param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
    offset = offset_param.AsDouble() if offset_param else 0
    
    absolute_elevation = level.Elevation + offset
    return level, offset, absolute_elevation

def get_most_common_level(floors):
    """Get the level that has the most floors in the selection."""
    if not floors:
        return None
    
    floor_levels = [get_floor_info(floor)[0] for floor in floors if get_floor_info(floor)[0]]
    level_counter = Counter(floor_levels)
    
    # Return the most common level or the highest level if there's a tie
    if level_counter:
        most_common = level_counter.most_common(1)[0][0]
        return most_common
    
    return None

def change_floor_level(floor, new_level):
    """Changes the level of a floor and adjusts offset to maintain position."""
    try:
        # Get current level, offset, and absolute elevation
        current_level, current_offset, current_absolute_elevation = get_floor_info(floor)
        if not current_level:
            return False
            
        # Find level parameter
        level_param = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM)
        offset_param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
        if not level_param or not offset_param:
            return False
            
        # Calculate new offset
        new_offset = current_absolute_elevation - new_level.Elevation
        
        # Change level first
        level_param.Set(new_level.Id)
        
        # Then adjust offset
        offset_param.Set(new_offset)
        
        return True
        
    except Exception as e:
        print("Error changing floor level: {}".format(str(e)))
        return False

def main():
    # Get selected floors
    selected_floors = get_selected_floors()
    
    if not selected_floors:
        forms.alert("Please select floor elements.", title="No Floors Selected")
        return
    
    # Get all levels
    all_levels = get_levels()
    level_options = {lvl.Name: lvl for lvl in all_levels}
    
    # Find the level that has the most selected elements
    most_common_level = get_most_common_level(selected_floors)
    default_level_name = most_common_level.Name if most_common_level else all_levels[0].Name
    
    # Let user select target level
    selected_level_name = forms.ask_for_one_item(
        sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
        default=default_level_name,
        prompt="Select target level:",
        title="Change Floor Level"
    )
    
    if not selected_level_name:
        return
    
    target_level = level_options[selected_level_name]
    
    # Process each floor
    success_count = 0
    fail_count = 0
    
    # Start a transaction
    with Transaction(doc, "Change Floor Level") as t:
        t.Start()
        
        try:
            for floor in selected_floors:
                if change_floor_level(floor, target_level):
                    success_count += 1
                else:
                    fail_count += 1
                    
            t.Commit()
        except Exception as e:
            t.RollBack()
            forms.alert("Error: {}".format(str(e)), title="Transaction Failed")
            return
    
    # Report results
    result_message = "{} floors Level changed to '{}'".format(success_count, target_level.Name)
    if fail_count > 0:
        result_message += "\n{} floors Level could not be changed".format(fail_count)
    
    forms.alert(result_message, title="Operation Complete")

if __name__ == "__main__":
    main()