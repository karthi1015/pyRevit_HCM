"""

Copyright (c) 2016 | Gui Talarico
Base script taken from pyRevit Respository.

pyRevit Notice
#################################################################
Copyright (c) 2014-2016 Ehsan Iran-Nejad
Python scripts for Autodesk Revit

This file is part of pyRevit repository at https://github.com/eirannejad/pyRevit

pyRevit is a free set of scripts for Autodesk Revit: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3, as published by
the Free Software Foundation.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

See this link for a copy of the GNU General Public License protecting this package.
https://github.com/eirannejad/pyRevit/blob/master/LICENSE
"""

__doc__ = 'Exports selected schedules as .txt and loads in Excel'

from Autodesk.Revit.DB import ViewSchedule, ViewScheduleExportOptions
from Autodesk.Revit.DB import ExportColumnHeaders, ExportTextQualifier
from Autodesk.Revit.DB import BuiltInCategory, ViewSchedule
from Autodesk.Revit.UI import TaskDialog

import os
import subprocess

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

desktop = os.path.expandvars('%temp%\\')
# desktop = os.path.expandvars('%userprofile%\\desktop')

vseop = ViewScheduleExportOptions()
# vseop.ColumnHeaders = ExportColumnHeaders.None
# vseop.TextQualifier = ExportTextQualifier.None
# vseop.FieldDelimiter = ','
# vseop.Title = False

selected_ids = uidoc.Selection.GetElementIds()

if not selected_ids.Count:
    '''If nothing is selected, use Active View'''
    selected_ids=[ doc.ActiveView.Id ]

for element_id in selected_ids:
    element = doc.GetElement(element_id)
    if not isinstance(element, ViewSchedule):
        print('No schedule in Selection. Skipping...')
        continue

    filename = "".join(x for x in element.ViewName if x not in ['*']) + '.csv'
    element.Export(desktop, filename, vseop)

    print('EXPORTED: {0}\n      TO: {1}\n'.format(element.ViewName, filename))
    EXCEL = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    if os.path.exists(EXCEL):
        print('Excel Found. Trying to open...')
        print('Filename is: ', filename)
        try:
            full_filepath = os.path.join(desktop, filename)
            os.system('start excel \"{path}\"'.format(path=full_filepath))
        except:
            print('Sorry, something failed:')
            print('Filepath: {}'.filename)
            print('EXCEL Path: {}'.format(EXCEL))
    else:
        print('Could not find excel. EXCEL: {}'.format(EXCEL))

print('Done')
