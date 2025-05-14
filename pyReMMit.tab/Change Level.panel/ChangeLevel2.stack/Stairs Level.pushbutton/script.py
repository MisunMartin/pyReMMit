"""Changes host level of selected stairs while maintaining their absolute position.
   Provides options for handling both bottom and top constraints."""
__title__ = 'Change Stair Level'
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

def get_selected_stairs():
    """Get selected stair elements from current selection."""
    selection = uidoc.Selection.GetElementIds()
    # Use category ID to identify stairs instead of class name
    stairs_category_id = ElementId(BuiltInCategory.OST_Stairs)
    return [doc.GetElement(elid) for elid in selection 
            if doc.GetElement(elid).Category 
            and doc.GetElement(elid).Category.Id == stairs_category_id]

def get_stair_bottom_info(stair):
    """Get base level, offset, and absolute elevation of a stair."""
    level_id = stair.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId()
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
        
    level = doc.GetElement(level_id)
    offset_param = stair.get_Parameter(BuiltInParameter.STAIRS_BASE_OFFSET)
    offset = offset_param.AsDouble() if offset_param else 0
    
    absolute_elevation = level.Elevation + offset
    return level, offset, absolute_elevation

def get_stair_top_info(stair):
    """Get top constraint level, offset, and absolute elevation of a stair."""
    level_id = stair.get_Parameter(BuiltInParameter.STAIRS_TOP_LEVEL_PARAM).AsElementId()
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
    
    top_level = doc.GetElement(level_id)
    if not top_level:
        return None, 0, 0
        
    top_offset_param = stair.get_Parameter(BuiltInParameter.STAIRS_TOP_OFFSET)
    top_offset = top_offset_param.AsDouble() if top_offset_param else 0
    
    absolute_top_elevation = top_level.Elevation + top_offset
    return top_level, top_offset, absolute_top_elevation

def get_most_common_level(stairs):
    """Get the level that has the most stairs in the selection."""
    if not stairs:
        return None
    
    stair_levels = [get_stair_bottom_info(stair)[0] for stair in stairs if get_stair_bottom_info(stair)[0]]
    level_counter = Counter(stair_levels)
    
    # Return the most common level or the highest level if there's a tie
    if level_counter:
        most_common = level_counter.most_common(1)[0][0]
        return most_common
    
    return None

def change_stair_bottom_level(stair, new_level):
    """Changes the Base Level of a stair and adjusts offset to maintain position."""
    try:
        # Get current level, offset, and absolute elevation
        current_level, current_offset, current_absolute_elevation = get_stair_bottom_info(stair)
        if not current_level:
            return False
            
        # Find level parameter
        level_param = stair.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
        offset_param = stair.get_Parameter(BuiltInParameter.STAIRS_BASE_OFFSET)
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
        print("Error changing stair Base Level: {}".format(str(e)))
        return False

def change_stair_top_level(stair, new_level):
    """Changes the Top Level of a stair and adjusts offset to maintain position."""
    try:
        # Get current top constraint info
        current_top_level, current_top_offset, current_top_absolute_elevation = get_stair_top_info(stair)
        if not current_top_level:
            return False
            
        # Get parameters needed to modify
        top_level_param = stair.get_Parameter(BuiltInParameter.STAIRS_TOP_LEVEL_PARAM)
        top_offset_param = stair.get_Parameter(BuiltInParameter.STAIRS_TOP_OFFSET)
        
        if not top_level_param or not top_offset_param:
            return False
        
        # Calculate new offset
        new_top_offset = current_top_absolute_elevation - new_level.Elevation
        
        # Set the top level parameter to the new level ID
        top_level_param.Set(new_level.Id)
        
        # Set the new offset
        top_offset_param.Set(new_top_offset)
        
        return True
        
    except Exception as e:
        print("Error changing stair Top Level: {}".format(str(e)))
        return False

def is_top_constrained_to_level(stair):
    """Check if stair's top is constrained to a level."""
    try:
        # Get the top level parameter
        top_level_param = stair.get_Parameter(BuiltInParameter.STAIRS_TOP_LEVEL_PARAM)
        if not top_level_param:
            return False
        
        # Check if it's connected to a level by seeing if the AsElementId is valid
        level_id = top_level_param.AsElementId()
        return level_id != ElementId.InvalidElementId
    except:
        return False

def main():
    # Get selected stairs
    selected_stairs = get_selected_stairs()
    
    if not selected_stairs:
        forms.alert("Please select stair elements.", title="No Stairs Selected")
        return
    
    # Get all levels
    all_levels = get_levels()
    level_options = {lvl.Name: lvl for lvl in all_levels}
    
    # Find the level that has the most selected elements
    most_common_level = get_most_common_level(selected_stairs)
    default_level_name = most_common_level.Name if most_common_level else all_levels[0].Name
    
    # Let user select target bottom constraint level
    selected_level_name = forms.ask_for_one_item(
        sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
        default=default_level_name,
        prompt="Select target Base Level:",
        title="Change Stair Base Level"
    )
    
    if not selected_level_name:
        return
    
    target_level = level_options[selected_level_name]
    
    # Ask user about top constraint handling
    top_constraint_options = ["Do not change Top Level", "Change Top Level"]
    selected_top_option = forms.ask_for_one_item(
        top_constraint_options,
        default=top_constraint_options[0],
        prompt="Top Level handling:",
        title="Top Level Options"
    )
    
    if not selected_top_option:
        return
    
    # If changing top constraint, ask for target level
    target_top_level = None
    if selected_top_option == "Change Top Level":
        selected_top_level_name = forms.ask_for_one_item(
            sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
            default=default_level_name,
            prompt="Select target Top Level:",
            title="Change Stair Top Level"
        )
        
        if not selected_top_level_name:
            return
            
        target_top_level = level_options[selected_top_level_name]
    
    # Process each stair
    bottom_success_count = 0
    bottom_fail_count = 0
    top_success_count = 0
    top_fail_count = 0
    top_not_applicable_count = 0
    
    # Start a transaction
    with Transaction(doc, "Change Stair Level") as t:
        t.Start()
        
        try:
            for stair in selected_stairs:
                # Change bottom constraint
                if change_stair_bottom_level(stair, target_level):
                    bottom_success_count += 1
                else:
                    bottom_fail_count += 1
                
                # Handle top constraint if requested
                if selected_top_option == "Change Top Level":
                    # Explicitly check if stair's top is constrained to level
                    is_top_level = is_top_constrained_to_level(stair)
                    
                    if is_top_level:
                        # Try to change top constraint
                        if change_stair_top_level(stair, target_top_level):
                            top_success_count += 1
                        else:
                            top_fail_count += 1
                    else:
                        top_not_applicable_count += 1
                    
            t.Commit()
        except Exception as e:
            t.RollBack()
            forms.alert("Error: {}".format(str(e)), title="Transaction Failed")
            return
    
    # Report results
    result_message = "{} stairs Base Level changed to \n'{}'".format(bottom_success_count, target_level.Name)
    if bottom_fail_count > 0:
        result_message += "\n{} stairs Base Level could not be changed!".format(bottom_fail_count)
    
    if selected_top_option == "Change Top Level":
        result_message += "\n\n{} stairs Top Level changed to \n'{}'".format(
            top_success_count, target_top_level.Name)
        if top_fail_count > 0:
            result_message += "\n{} stairs Top Level could not be changed!".format(top_fail_count)
        if top_not_applicable_count > 0:
            result_message += "\n\n{} stairs did not have a Top Level constraint".format(top_not_applicable_count)
    
    forms.alert(result_message, title="Operation Complete")

if __name__ == "__main__":
    main()