# -*- coding: utf-8 -*-
#crash test2
# IMPORT
import clr
import System 
import rpw
from rpw import revit, db, ui, DB, UI
from rpw.ui.forms import *
from System import Array
from System.Runtime.InteropServices import Marshal
import sys
import Autodesk
# sys.path.append(r"C:\\Program Files (x86)\\Python37-32\\Lib")
# from tkinter import *
from pyrevit import script

# e_a = str("\xe9")
# a_a = str("\xe0")
# Fenetre = Tk()
# bouton = Button(Fenetre, text = "quitter", command=Fenetre.destroy)
# bouton.pack()
# Fenetre.mainloop()





output = script.get_output()

output.set_height(600)
output.get_title()
output.set_title()





# # USERFORM BIM CHECKER
# components = [Label('Check a choisir'),
#               CheckBox('checkbox1', 'Verifier le nom de la maquette'),
#               CheckBox('checkbox2', 'Verifier si la maquette est centrale ou local'),
#               CheckBox('checkbox3', 'Verifier le nombre de groupe dans la maquette' ),
#               CheckBox('checkbox4', 'Verifier emplacement partage'),
#               CheckBox('checkbox5', 'Verifier le point de base'),
#               CheckBox('checkbox6', 'Verifier les niveaux'),
#               CheckBox('checkbox7', 'Verifier les quadrillages'),
#               CheckBox('checkbox8', 'Verifier les avertissements'),
#               CheckBox('checkbox9', 'Verifier si chaque vue est utilisee pour une vue en plan'),
#               CheckBox('checkbox10', 'Verifier le poids de la maquette'),
#               CheckBox('checkbox11', 'Verifier la version Revit'),
#               Label('Version Revit'),
#               ComboBox('combobox1', {'2015':2015, '2016':2016, '2017':2017, '2018':2018, '2019':2019}),
#               Label('Nom maquette exacte'),
#               TextBox('textbox1', Text="EXEMPLE_MAQUETTE.rvt"),
#               Label('Poids maximum maquette'),
#               ComboBox('combobox2', {'150':150, '200':200, '250':250, '300':300, '350':350, '400':400, '450':450}),
#               Label('choisir maquette QNP'),
#               Button(button_text='QNP'),             
#               Label('choisir visa Excel'),
#               Button(button_text='VISA BIM'),
#               Separator(), 
#               Button('OK GO'),]
# form = FlexForm('BIM CHECKER DU FUTUR', components)
# form.show()










































