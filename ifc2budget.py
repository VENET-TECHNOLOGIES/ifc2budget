import multiprocessing
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util as util
import PySimpleGUI as sg
import re
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


sg.theme("DarkTeal2")
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
                ifcTree.Insert("", story.get_info()["id"], story.get_info()["Name"], [])
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
                workbook = xlsxwriter.Workbook('ifc.xlsx')
                worksheet = workbook.add_worksheet()
                
                for story in stories_list:
                    for element in story["decomposition"]:
                        element_list = []
                        #print(element.get_info())
                        e_type = element.get_info()["type"]
                        e_name = element.get_info()["Name"]
                        e_name_clean = e_name.split(':')[:-1]

                        if "Ceiling" in e_name_clean:
                            e_name_clean.remove("Ceiling")
                        if "Basic Wall" in e_name_clean:
                            e_name_clean.remove("Basic Wall")    
                        if "Floor" in e_name_clean:
                            e_name_clean.remove("Floor")
                        if "Compound Ceiling" in e_name_clean:
                            e_name_clean.remove("Compound Ceiling")
                        if "Railing" in e_name_clean:
                            e_name_clean.remove("Railing")
                        if "Cast-In-Place Stair" in e_name_clean:
                            e_name_clean.remove("Cast-In-Place Stair")
                        if "Ramp" in e_name_clean:
                            e_name_clean.remove("Ramp")

                        e_name_clean = (':').join(e_name_clean)

                        try:
                            e_psets = util.element.get_psets(element)
                            keys= list(e_psets.keys())
                            for key in keys:
                                has_quantities = key.find("Quantities")
                                if has_quantities > 0:
                                    print(e_name_clean, "-", e_type, "-", key, "-", e_psets[key])
                        except Exception as error:
                            open_modal("Error", "No measurements:" + error)

                workbook.close()
            else:
                open_modal("Error", "No IFC loaded!")

        except Exception as error:
            print("Error building the excel file! -- ", error)
            

        