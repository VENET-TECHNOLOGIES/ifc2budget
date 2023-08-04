def clean_ifc_element(element):
    e_type = element.get_info()["type"]
    e_name = element.get_info()["Name"]

    if e_name != None:
        #print("Pre-split: ", element, e_type, e_name)
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
        return(e_name_clean,e_type)
    else:
        return("Sin nombre!", e_type)
    