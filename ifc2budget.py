import multiprocessing
import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util as util
import PySimpleGUI as sg
from pprint import pprint

ifcTree = sg.TreeData()

sgTree=sg.Tree(data=ifcTree,
   headings=["Type", "Name"],
   auto_size_columns=True,
   select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
   col0_width=2,
   key='-TREE-',
   show_expanded=False,
   enable_events=True,
   expand_x=True,
   expand_y=True,
)

sg.theme("DarkTeal2")
layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(), sg.FileBrowse(key="-IN-")],[sg.Button("Read!")],[sgTree]]
settings = ifcopenshell.geom.settings()


window =  sg.Window('IFC2Budget', layout, size=(600,150))

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
                print(story.get_info())
                print(story.get_info()["type"])
                print(story.get_info()["Name"])
                print(story.get_info()["id"])
                ifcTree.Insert("Test Type", "Test Name", [12,13])
                story_obj = {}
                story_obj["story"] = story

                decomposition = util.element.get_decomposition(story)
                story_obj["decomposition"] = []
                story_obj["spaces"] = []
                for element in decomposition:
                    if element.get_info()['type'] == 'IfcSpace':
                        new_space = {}
                        new_space["space"] = element
                        new_space["space_decomposition"] = util.element.get_decomposition(element)
                        story_obj["spaces"].append(new_space)
                    else:
                        story_obj["decomposition"].append(element)
                stories_list.append(story_obj)
                

            #pprint(stories_list)
            #pprint(site_elements)
            # iterator = ifcopenshell.geom.iterator(settings, model, multiprocessing.cpu_count())
            # if iterator.initialize():
            #     while True:
            #         element = iterator.get_native()
            #         object = iterator.get()

            #         print(object.product)
            #         tree.add_element(element)
            #         if not iterator.next():
            #             break

        except Exception as error:
            print("Error reading the file! -- ", error)