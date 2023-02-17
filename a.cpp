# import pandas lib as pd
import pandas as pd
import math
# import module
from bs4 import BeautifulSoup
from datetime import datetime

def convert_excel_html():
    # read by default 1st sheet of an excel file
    dataframe = pd.read_excel('..\\niftarim.xlsx')
    # assign URL
    index_content = open("index.html", "r", encoding="utf-8")
    s = BeautifulSoup(index_content,"lxml")

    for index in range(dataframe.shape[0]):
        new_tr = s.new_tag("tr")
        new_td = s.new_tag('td')
        image_td = s.new_tag('td')
        new_bold = s.new_tag('b')
        new_p = s.new_tag('p')
        name_p = s.new_tag('p')
        new_desc = s.new_tag('p')
        new_grave = s.new_tag('p')
        new_img = s.new_tag('img')
        new_img["src"] = dataframe.Image[index]
        new_img["width"] = 300
        new_img["height"] = 350
        new_bold.string = dataframe.Name[index] + " " + dataframe.FamilyName[index]
        born_string = "נולד בתאריך "
        born_year_string = "נולד בשנת "
        place_string = ""
        die_string = " ונפטר בתאריך "
        die_year_string = " ונפטר בשנת "
        grave_start_string = "נטמן בחלקה "
        age_string = ""
        if dataframe.Gender[index] == "F":
            born_string = "נולדה בתאריך "
            born_year_string = "נולדה בשנת "
            die_string = " ונפטרה בתאריך "
            die_year_string = " ונפטרה בשנת "
            grave_start_string = "נטמנה בחלקה "
        if dataframe.PassAwayAge[index] > 1:
            age_string = str(int(dataframe.PassAwayAge[index]))
        elif dataframe.PassAwayAge[index] < 1:
            age_string = str(int(dataframe.PassAwayAge[index]*12)) + " חודשים"
        else:
            age_string = "שנה"
        if dataframe.BirthLocation[index] != "-":
            place_string = " ב"+ dataframe.BirthLocation[index]
        if type(dataframe.BirthDate[index]) == datetime and type(dataframe.PassAwayDate[index]) == datetime:
            new_p.string = born_string + "{0}/{1}/{2}".format(dataframe.BirthDate[index].day, dataframe.BirthDate[index].month, dataframe.BirthDate[index].year) + place_string + \
                           die_string + "{0}/{1}/{2}".format(dataframe.PassAwayDate[index].day, dataframe.PassAwayDate[index].month, dataframe.PassAwayDate[index].year) + "  בגיל " \
                           + age_string + "."
        elif type(dataframe.PassAwayDate[index]) == datetime:
            new_p.string = born_year_string + str(dataframe.BirthDate[index]) + \
                           die_string + "{0}/{1}/{2}".format(dataframe.PassAwayDate[index].day,
                                                             dataframe.PassAwayDate[index].month,
                                                             dataframe.PassAwayDate[index].year) + "  בגיל " \
                           + age_string + "."
        else:
            new_p.string = born_year_string + str(dataframe.BirthDate[index]) + \
                           die_year_string + str(dataframe.PassAwayDate[index]) + "  בגיל " \
                           + age_string + "."
        if dataframe.Death[index] != "-":
            new_p.string = born_string + "{0}/{1}/{2}".format(dataframe.BirthDate[index].day,
                                                              dataframe.BirthDate[index].month,
                                                              dataframe.BirthDate[index].year) + place_string + "." \
                           +" " + dataframe.Death[index] + " בתאריך " + "{0}/{1}/{2}".format(dataframe.PassAwayDate[index].day,
                                                             dataframe.PassAwayDate[index].month,
                                                             dataframe.PassAwayDate[index].year) + ".  בן " \
                           + age_string + " בנפלו."
        print(new_bold.string)
        if dataframe.Description[index] == "-":
            new_desc.string = ""
        else:
            new_desc.string = dataframe.Description[index]
        if type(dataframe.GraveArea[index]) == str:
            new_grave.string = grave_start_string + str(dataframe.GraveArea[index]) + ", שורה " + str(int(dataframe.GraveRow[index])) + ", קבר " +str(int(dataframe.GraveNumber[index])) + "."
        name_p.append(new_bold)
        new_td.append(name_p)

        new_td.append(new_p)
        new_td.append(new_desc)
        new_td.append(new_grave)

        image_td.append(new_img)
        new_tr.append(new_td)
        new_tr.append(image_td)
        s.table.append(new_tr)

    with open("output1.html", "w", encoding="utf-8") as file:
        file.write(str(s))

convert_excel_html()
pass
