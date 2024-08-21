# -*- coding: utf-8 -*-
#region import libary
import clr
import os
import sys
import csv
import io

clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")
import RevitServices
import Autodesk
import System
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import *

from System.Windows import MessageBox
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader
#endregion
#region get infor
dir_path = os.path.dirname(os.path.realpath(__file__))
xaml_file_path = os.path.join(dir_path, "Window.xaml")
#revit infor
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
activeView = doc.ActiveView
#endregion


class Utils:
    def __init__(self):
        self.symbol_list = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()

    def get_all_family_names (self):
        return sorted(set([symbol.FamilyName for symbol in self.symbol_list]))
    
    def get_all_type_names (self, family_name):
        return sorted([symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                       for symbol in self.symbol_list if symbol.FamilyName == family_name])
    
    def get_title_block_id (self, family_name, type_name):
        for type in self.symbol_list:
            if type.FamilyName == family_name and type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == type_name:
                return type.Id
        return None
    
    def set_parameters (self, sheet, parameter_name, parameter_value):
        params = sheet.LookupParameter(parameter_name)
        if params is not None and params.IsReadOnly == False:
            if params.StorageType == StorageType.Double:
                params.Set(float(parameter_value))
            elif params.StorageType == StorageType.Integer:
                params.Set(int(parameter_value))
            else: params.Set(parameter_value)

class WPFWindow:
    def __init__(self):
        pass

    def load_window (self):
        #import window from .xaml file path
        file_stream = FileStream(xaml_file_path, FileMode.Open, FileAccess.Read)
        window = XamlReader.Load(file_stream)

        #controls
        self.cbb_typename = window.FindName("cbb_TypeName")
        self.tb_Directory = window.FindName("tb_Directory")
        self.cbb_family = window.FindName("cbb_Family")
        self.bt_Cancel = window.FindName("bt_Cancel")
        self.bt_Browse = window.FindName("bt_Browse")
        self.bt_Create = window.FindName("bt_Create")

        #bindingdata
        self.bindind_data()
        self.window = window
        return window
    
    def bindind_data (self):
        families_names = Utils().get_all_family_names()
        type_names = Utils().get_all_type_names(families_names[0])

        self.cbb_family.ItemsSource = families_names
        self.cbb_typename.ItemsSource = type_names
        self.cbb_family.SelectedIndex = 0
        self.cbb_typename.SelectedIndex = 0

        self.cbb_family.SelectionChanged += self.cbb_Family_SelectionChanged
        self.bt_Cancel.Click += self.Cancel_Click
        self.bt_Browse.Click += self.Browse_Click
        self.bt_Create.Click += self.Create_Click

    def cbb_Family_SelectionChanged(self, sender, e):
        try:
            selected_family = self.cbb_family.SelectedValue
            self.cbb_typename.ItemsSource = Utils().get_all_type_names(selected_family)
            self.cbb_typename.SelectedIndex = 0
        except:
            pass

    def Create_Click(self, sender, e):
        #get existing sheet numbers
        all_sheet = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
        existing_sheet_numbers = []
        for sheet in all_sheet:
            if sheet.IsPlaceholder == False:
                existing_sheet_numbers.append(sheet.SheetNumber)

        total_existing_sheet_numbers = len(existing_sheet_numbers)

        #get title Block id
        block_id = Utils().get_title_block_id(self.cbb_family.SelectedItem, self.cbb_typename.SelectedItem)

        #get data rows
        file_path = self.tb_Directory.Text
        if file_path == "":
            MessageBox.Show("Select a .csv file or copy/paste the file path!", "Message")
        else:
            data_rows = []
            with io.open(file_path, mode='r', encoding='utf-8') as csv_data:
                data_reader = csv.reader(csv_data, delimiter = ',')
                for row in data_reader:
                    data_rows.append(row)

            #get first row and delete first row
            row_0 = data_rows[0]
            del data_rows[0]

            #create sheets
            t = Transaction(doc, "create sheets")
            t.Start()
            for i in range(len(data_rows)):
                if block_id is not None:
                    sheet = ViewSheet.Create(doc, block_id)
                    try:
                        for j in range(len(row_0)):
                            para_name = row_0[j]
                            para_value = data_rows[i][j]
                            if para_name == "Sheet Number" and list(existing_sheet_numbers).__contains__(para_value):
                                doc.Delete(sheet.Id)
                            else: Utils().set_parameters(sheet, para_name, para_value)
                    except:
                        try:
                             doc.Delete(sheet.Id)
                        except:
                            pass
            #update model
            doc.Regenerate() 
            t.Commit()

            #show message
            total_sheet = len(list(FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType()))
            message = str(total_sheet - total_existing_sheet_numbers) + " sheets created!"
            MessageBox.Show(message, "Message")
            self.window.Close()
    
    def Cancel_Click (self, sender, e):
        self.window.Close()
    
    def Browse_Click (self, sender, e):
        dlg = OpenFileDialog()
        dlg.Filter = "CSV files (*.csv)|*.csv"
        if dlg.ShowDialog() == DialogResult.OK:
            self.tb_Directory.Text = dlg.FileName

class Main:
    def main_task(self):
        try:
            window = WPFWindow().load_window()
            window.ShowDialog()
        except Exception as e:
            MessageBox.Show(str(e), "Message")

if __name__ == "__main__":
    Main ().main_task()





