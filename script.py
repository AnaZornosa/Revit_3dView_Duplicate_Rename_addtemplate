# -*- coding: utf-8 -*-
__title__ = ("Duplicate"
             "/Rename/"
             "Template")
__doc__ = """Version = 1.0
Date    = 9.15.2024
_____________________________________________________________________
Description:
Duplicate and rename 3D views, then assign view templates based on selected system.
_____________________________________________________________________
How-to:
-> Select views in the project browser
-> Choose which system (Architectural, Mechanical, Plumbing, Fire Protection, Electrical)
-> Views will be duplicated, renamed, and templates will be applied.
_____________________________________________________________________
Last update:
- [09.15.2024] - 1.0 RELEASE
_____________________________________________________________________
Author: Ana Zornosa
"""

# IMPORTS, REVIT DATA BASE, PYREVIT FORMS AND REVIT IN C#
from Autodesk.Revit.DB import *
from pyrevit import revit, forms
import clr

clr.AddReference("System")
from System.Collections.Generic import List

# VARIABLES
doc = __revit__.ActiveUIDocument.Document  # Active Revit document


# MAIN
def main():
    # Select 3D views from the project browser
    sel_views = forms.select_views()

    # Ensure Views are selected
    if not sel_views:
        forms.alert('No Views Selected', exitscript=True)

    # Ask the user to choose which system they want to duplicate views for
    options = ['Architectural', 'Mechanical', 'Plumbing', 'Fire Protection', 'Electrical']
    system = forms.ask_for_one_item(options, title='Choose System for Duplication')

    if not system:
        forms.alert('No system selected', exitscript=True)

    # Map system choices to prefixes and view templates (all in capital letters)
    system_mapping = {
        'Architectural': {'prefix': 'A_', 'template': 'ARCHITECTURE GLUE'},
        'Mechanical': {'prefix': 'MD_', 'template': 'MECHANICAL GLUE'},
        'Plumbing': {'prefix': 'PL_', 'template': 'PLUMBING GLUE'},
        'Fire Protection': {'prefix': 'FP_', 'template': 'FIRE PROTECTION GLUE'},
        'Electrical': {'prefix': 'EE_', 'template': 'ELECTRICAL GLUE'}
    }

    # Retrieve the prefix and template based on the selected system
    prefix = system_mapping[system]['prefix']
    template_name = system_mapping[system]['template']

    # Start a transaction to make changes
    t = Transaction(doc, 'Duplicate and Rename 3D Views')
    t.Start()

    duplicated_count = 0  # Counter for duplicated views

    try:
        for view in sel_views:
            if isinstance(view, View3D):
                # Duplicate the view
                duplicated_view_id = view.Duplicate(ViewDuplicateOption.Duplicate)
                duplicated_view = doc.GetElement(duplicated_view_id)

                # Rename the duplicated view
                old_name = view.Name
                underscore_index = old_name.find('_')  # Find where the FIRST
                # Make sure to remove any extra underscores and add exactly one underscore
                if underscore_index != -1:  # If there is an underscore in the old name
                    new_name = prefix + old_name[underscore_index + 1:]  # Add only the part after the underscore
                else:
                    new_name = prefix + old_name  # If no underscore, just add the prefix

                # Ensure unique name
                for i in range(20):
                    try:
                        duplicated_view.Name = new_name
                        print('{} -> {}'.format(old_name, new_name))
                        break
                    except:
                        new_name += '*'

                # Find the view template by name
                template_collector = FilteredElementCollector(doc).OfClass(View)
                new_template = None

                # Loop through the collector to find the template by name
                for template in template_collector:
                    if template.Name == template_name:  # Ensure exact match
                        new_template = template
                        break

                # Assign the correct view template if found
                if new_template:
                    duplicated_view.ViewTemplateId = new_template.Id
                else:
                    # Alert if the template was not found
                    print("Template '{}' not found for '{}'".format(template_name, new_name))

                duplicated_count += 1

        t.Commit()  # Commit the transaction

        # Show success message
        forms.alert('{} views duplicated, renamed, and template applied.'.format(duplicated_count))

    except Exception as e:
        # Rollback transaction if an error occurs
        t.RollBack()
        forms.alert('An error occurred: {}'.format(e))


# Run the main function
main()