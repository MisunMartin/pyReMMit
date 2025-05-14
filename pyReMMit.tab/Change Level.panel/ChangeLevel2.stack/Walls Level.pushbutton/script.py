"""Changes host level of selected walls while maintaining their absolute position.
   Provides options for handling both bottom and top constraints."""
__title__ = 'Change Wall Level'
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

def get_selected_walls():
    """Get selected wall elements from current selection."""
    selection = uidoc.Selection.GetElementIds()
    return [doc.GetElement(elid) for elid in selection if isinstance(doc.GetElement(elid), Wall)]

def get_wall_bottom_info(wall):
    """Get base level, offset, and absolute elevation of a wall."""
    level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
    if level_id == ElementId.InvalidElementId:
        return None, 0, 0
        
    level = doc.GetElement(level_id)
    offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    offset = offset_param.AsDouble() if offset_param else 0
    
    absolute_elevation = level.Elevation + offset
    return level, offset, absolute_elevation

def get_wall_top_info(wall):
    """Get top constraint level, offset, and absolute elevation of a wall."""
    # Get the height type parameter to determine if top is constrained to level
    height_type_param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
    if not height_type_param:
        return None, 0, 0
    
    # Check if it's constrained to a level by checking if height_type_param.AsElementId() is valid
    level_id = height_type_param.AsElementId()
    
    # If the ElementId is valid, it's connected to a level
    if level_id != ElementId.InvalidElementId:
        top_level = doc.GetElement(level_id)
        if not top_level:
            return None, 0, 0
            
        top_offset_param = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        top_offset = top_offset_param.AsDouble() if top_offset_param else 0
        
        absolute_top_elevation = top_level.Elevation + top_offset
        return top_level, top_offset, absolute_top_elevation
    else:
        # Wall is not connected to a level at top
        return None, 0, 0

def get_most_common_level(walls):
    """Get the level that has the most walls in the selection."""
    if not walls:
        return None
    
    wall_levels = [get_wall_bottom_info(wall)[0] for wall in walls if get_wall_bottom_info(wall)[0]]
    level_counter = Counter(wall_levels)
    
    # Return the most common level or the highest level if there's a tie
    if level_counter:
        most_common = level_counter.most_common(1)[0][0]
        return most_common
    
    return None

def change_wall_bottom_level(wall, new_level):
    """Changes the base constraint level of a wall and adjusts offset to maintain position."""
    try:
        # Get current level, offset, and absolute elevation
        current_level, current_offset, current_absolute_elevation = get_wall_bottom_info(wall)
        if not current_level:
            return False
            
        # Find level parameter
        level_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
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
        print("Error changing wall base constraint: {}".format(str(e)))
        return False

def change_wall_top_level(wall, new_level):
    """Changes the top constraint level of a wall and adjusts offset to maintain position."""
    try:
        # Get current top constraint info
        current_top_level, current_top_offset, current_top_absolute_elevation = get_wall_top_info(wall)
        if not current_top_level:
            return False
            
        # Get parameters needed to modify
        height_type_param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
        top_offset_param = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        
        if not height_type_param or not top_offset_param:
            return False
        
        # Calculate new offset
        new_top_offset = current_top_absolute_elevation - new_level.Elevation
        
        # Set the height type parameter to the new level ID
        height_type_param.Set(new_level.Id)
        
        # Set the new offset
        top_offset_param.Set(new_top_offset)
        
        return True
        
    except Exception as e:
        print("Error changing wall top constraint: {}".format(str(e)))
        return False

def is_top_constrained_to_level(wall):
    """Check if wall's top is constrained to a level using ID comparison."""
    try:
        # Get the height type parameter
        height_type_param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
        if not height_type_param:
            return False
        
        # Check if it's connected to a level by seeing if the AsElementId is valid
        level_id = height_type_param.AsElementId()
        return level_id != ElementId.InvalidElementId
    except:
        return False

def main():
    # Get selected walls
    selected_walls = get_selected_walls()
    
    if not selected_walls:
        forms.alert("Please select wall elements.", title="No Walls Selected")
        return
    
    # Get all levels
    all_levels = get_levels()
    level_options = {lvl.Name: lvl for lvl in all_levels}
    
    # Find the level that has the most selected elements
    most_common_level = get_most_common_level(selected_walls)
    default_level_name = most_common_level.Name if most_common_level else all_levels[0].Name
    
    # Let user select target bottom constraint level
    selected_level_name = forms.ask_for_one_item(
        sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
        default=default_level_name,
        prompt="Select target base constraint level:",
        title="Change Wall Base Constraint"
    )
    
    if not selected_level_name:
        return
    
    target_level = level_options[selected_level_name]
    
    # Ask user about top constraint handling
    top_constraint_options = ["Do not change top constraint", "Change top constraint"]
    selected_top_option = forms.ask_for_one_item(
        top_constraint_options,
        default=top_constraint_options[0],
        prompt="Top constraint handling:",
        title="Top Constraint Options"
    )
    
    if not selected_top_option:
        return
    
    # If changing top constraint, ask for target level
    target_top_level = None
    if selected_top_option == "Change top constraint":
        selected_top_level_name = forms.ask_for_one_item(
            sorted(level_options.keys(), key=lambda name: level_options[name].Elevation, reverse=True),
            default=default_level_name,
            prompt="Select target top constraint level:",
            title="Change Wall Top Constraint"
        )
        
        if not selected_top_level_name:
            return
            
        target_top_level = level_options[selected_top_level_name]
    
    # Process each wall
    bottom_success_count = 0
    bottom_fail_count = 0
    top_success_count = 0
    top_fail_count = 0
    top_not_applicable_count = 0
    
    # Start a transaction
    with Transaction(doc, "Change Wall Level") as t:
        t.Start()
        
        try:
            for wall in selected_walls:
                # Change bottom constraint
                if change_wall_bottom_level(wall, target_level):
                    bottom_success_count += 1
                else:
                    bottom_fail_count += 1
                
                # Handle top constraint if requested
                if selected_top_option == "Change top constraint":
                    # Explicitly check if wall's top is constrained to level
                    is_top_level = is_top_constrained_to_level(wall)
                    
                    if is_top_level:
                        # Try to change top constraint
                        if change_wall_top_level(wall, target_top_level):
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
    result_message = "{} walls Base Constraint changed to \n'{}'".format(bottom_success_count, target_level.Name)
    if bottom_fail_count > 0:
        result_message += "\n{} walls Base Constraint could not be changed!".format(bottom_fail_count)
    
    if selected_top_option == "Change top constraint":
        result_message += "\n\n{} walls Top Constraint changed to \n'{}'".format(
            top_success_count, target_top_level.Name)
        if top_fail_count > 0:
            result_message += "\n{} walls Top Constraint could not be changed!".format(top_fail_count)
        if top_not_applicable_count > 0:
            result_message += "\n\n{} walls were set to 'Unconnected'".format(top_not_applicable_count)
    
    forms.alert(result_message, title="Operation Complete")

if __name__ == "__main__":
    main()