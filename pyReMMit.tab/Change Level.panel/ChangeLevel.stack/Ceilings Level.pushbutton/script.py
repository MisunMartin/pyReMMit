"""Changes host level of selected ceilings while maintaining their absolute position."""
__title__ = 'Change Ceiling Level'
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

def get_selected_ceilings():
    """Get selected ceiling elements from current selection."""
    selection = uidoc.Selection.GetElementIds()
    return [doc.GetElement(elid) for elid in selection if isinstance(doc.GetElement(elid), Ceiling)]

def get_ceiling_info(ceiling):
    """Get level, offset, and absolute elevation of a ceiling."""
    level_id = ceiling.LevelId
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
        
    level = doc.GetElement(level_id)
    offset_param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
    offset = offset_param.AsDouble() if offset_param else 0
    
    absolute_elevation = level.Elevation + offset
    return level, offset, absolute_elevation

def get_most_common_level(ceilings):
    """Get the level that has the most ceilings in the selection."""
    if not ceilings:
        return None
    
    ceiling_levels = [get_ceiling_info(ceiling)[0] for ceiling in ceilings if get_ceiling_info(ceiling)[0]]
    level_counter = Counter(ceiling_levels)
    
    # Return the most common level or the highest level if there's a tie
    if level_counter:
        most_common = level_counter.most_common(1)[0][0]
        return most_common
    
    return None

def change_ceiling_level(ceiling, new_level):
    """Changes the level of a ceiling and adjusts offset to maintain position."""
    try:
        # Get current level, offset, and absolute elevation
        current_level, current_offset, current_absolute_elevation = get_ceiling_info(ceiling)
        if not current_level:
            return False
            
        # Find level parameter
        level_param = ceiling.get_Parameter(BuiltInParameter.LEVEL_PARAM)
        offset_param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
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
        print("Error changing ceiling level: {}".format(str(e)))
        return False

def main():
    # Get selected ceilings
    selected_ceilings = get_selected_ceilings()
    
    if not selected_ceilings:
        forms.alert("Please select ceiling elements.", title="No Ceilings Selected")
        return
    
    # Get all levels
    all_levels = get_levels()
    level_options = {lvl.Name: lvl for lvl in all_levels}
    
    # Find the level that has the most selected elements
    most_common_level = get_most_common_level(selected_ceilings)
    default_level_name = most_common_level.Name if most_common_level else all_levels[0].Name
    
    # Let user select target level
    selected_level_name = forms.ask_for_one_item(
        sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
        default=default_level_name,
        prompt="Select target level:",
        title="Change Ceiling Level"
    )
    
    if not selected_level_name:
        return
    
    target_level = level_options[selected_level_name]
    
    # Process each ceiling
    success_count = 0
    fail_count = 0
    
    # Start a transaction
    with Transaction(doc, "Change Ceiling Level") as t:
        t.Start()
        
        try:    
            for ceiling in selected_ceilings:
                if change_ceiling_level(ceiling, target_level):
                    success_count += 1
                else:
                    fail_count += 1
                    
            t.Commit()
        except Exception as e:
            t.RollBack()
            forms.alert("Error: {}".format(str(e)), title="Transaction Failed")
            return
    
    # Report results
    result_message = "{} ceilings Level changed to '{}'".format(success_count, target_level.Name)
    if fail_count > 0:
        result_message += "\n{} ceilings Level could not be changed".format(fail_count)
    
    forms.alert(result_message, title="Operation Complete")

if __name__ == "__main__":
    main()