


def get_name_array(name):
    if (len(name) < 20):
        return [name]
    
    nameArr = name.split(" ")

    if (len(nameArr) == 1):
        return nameArr
    
    # if (len(nameArr) > 3):
    return nameArr[:3] 





