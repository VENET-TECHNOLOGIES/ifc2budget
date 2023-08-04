import multiprocessing
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util as util
import PySimpleGUI as sg
import re
import utils
import xlsxwriter

from pprint import pprint


def open_modal(windowTitle, message):
    layout = [[sg.Text(message, key="new")]]
    window = sg.Window(windowTitle, layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        
    window.close()

id_set = ["ID", "Name", "Story", "Space"]
measurements_set = ["Width", "Height", "Length", "NetArea", "NetSideArea" "OuterSurfaceArea", "GrossSurfaceArea", "GrossFootprintArea", "NetVolume", "NetWeight"]

col_dict = {"ID":                   0, 
            "Name":                 1, 
            "Story":                2,
            "Space":                3,
            "Width":                4,
            "Height":               5,
            "Length":               6,
            "NetArea":              7,
            "NetSideArea":          8,
            "OuterSurfaceArea":     9,
            "GrossSurfaceArea":     10,
            "GrossFootprintArea":   11,
            "NetVolume":            12,
            "NetWeight":            13,
            }
inv_col = {v: k for k, v in col_dict.items()}

ifc_is_read = False
ifcTree = sg.TreeData()

sgTree=sg.Tree(data=ifcTree,
                headings=["Type", "Name"],
                auto_size_columns=True,
                select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                col0_width=3,
                col0_heading='Story',
                key='-TREE-',
                show_expanded=False,
                enable_events=True,
                expand_x=True,
                expand_y=True,
                )


sg.theme("Default1")
layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(), sg.FileBrowse(key="-IN-"), sg.Button("Read!")],[],[sgTree],[sg.Button("ToExcel!")]]
settings = ifcopenshell.geom.settings()

window =  sg.Window('IFC2Budget', layout, size=(800,600))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Read!":
        try:
            print(values["-IN-"])
            model = ifcopenshell.open(values["-IN-"])
            building = model.by_type("IfcBuilding")
            stories = model.by_type("IfcBuildingStorey")
            stories_list = []
            for story in stories:                
                story_obj = {}
                story_obj["story"] = story

                decomposition = util.element.get_decomposition(story)
                
                story_obj["decomposition"] = []
                story_obj["spaces"] = []
                story_name = story.get_info()["Name"]
                story_id = story.get_info()["id"]
                story_obj["name"] = story_name
                story_obj["id"] = story_id

                ifcTree.Insert("", story_id, story_name, [])
                for element in decomposition:
                    if element.get_info()['type'] == 'IfcSpace':
                        new_space = {}
                        new_space["space"] = element
                        new_space_longname = new_space["space"].get_info()["LongName"]
                        new_space["space_decomposition"] = util.element.get_decomposition(element)
                        story_obj["spaces"].append(new_space)
                        ifcTree.Insert(story.get_info()["id"], new_space_longname, new_space_longname, [])
                    else:
                        story_obj["decomposition"].append(element)                        
                stories_list.append(story_obj)
            
            window['-TREE-'].update(ifcTree)
            ifc_is_read = True
        except Exception as error:
            print("Error reading the file! -- ", error)

    elif event == "ToExcel!":
        try:
            if ifc_is_read:
                print("To Excel")
                row = 0

                workbook = xlsxwriter.Workbook('ifc.xlsx')
                worksheet = workbook.add_worksheet()
                
                #Setting first row
                for idx, row_name in enumerate(inv_col):
                    worksheet.write(row,idx, inv_col[idx])
                row += 1
                
                for story in stories_list:
                    print("="*20, story["name"], "="*20)
                    for element in story["decomposition"]:
                        element_list = []
                        #print(element.get_info())
                        e_name, e_type = utils.clean_ifc_element(element)
                        try:
                            e_psets = util.element.get_psets(element)
                            keys= list(e_psets.keys())
                            for key in keys:
                                has_quantities = key.find("Quantities")
                                e_quantities = e_psets[key]

                                #This, to excel
                                if has_quantities > 0:
                                    print(e_name, "-", e_type, "-", key, "-", e_quantities)
                                    worksheet.write(row, col_dict["ID"], e_quantities["id"])
                                    worksheet.write(row, col_dict["Name"], e_name)
                                    worksheet.write(row, col_dict["Story"], story["name"])
                                    for measure in measurements_set:
                                        if measure in e_quantities:
                                            worksheet.write(row, col_dict[measure], e_quantities[measure])
                                    row += 1

                        except Exception as error:
                            open_modal("Error", "No measurements:" + error)
                        
                    for space in story["spaces"]:
                        space_name = space["space"].get_info()["LongName"]
                        print("*"*20, space_name, "*"*20)
                        for element in space["space_decomposition"]:
                            element_list = []
                            #print(element.get_info())
                            e_name, e_type = utils.clean_ifc_element(element)
                            try:
                                e_psets = util.element.get_psets(element)
                                keys= list(e_psets.keys())
                                for key in keys:
                                    has_quantities = key.find("Quantities")
                                    #This, to excel
                                    if has_quantities > 0:
                                        print(e_name, "-", e_type, "-", key, "-", e_quantities)
                                        worksheet.write(row, col_dict["ID"], e_quantities["id"])
                                        worksheet.write(row, col_dict["Name"], e_name)
                                        worksheet.write(row, col_dict["Story"], story["name"])
                                        worksheet.write(row, col_dict["Space"], space_name)
                                        for measure in measurements_set:
                                            if measure in e_quantities:
                                                worksheet.write(row, col_dict[measure], e_quantities[measure])
                                        row += 1
                            except Exception as error:
                                open_modal("Error", "No measurements:" + error)

                workbook.close()
            else:
                open_modal("Error", "No IFC loaded!")

        except Exception as error:
            print("Error building the excel file! -- ", error)
            

        