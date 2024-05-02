import clr
import re
import sys
import os
import shutil
import time
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application, Form, Button, ListBox, SelectionMode, AnchorStyles, TextBox, OpenFileDialog, DialogResult, SaveFileDialog, Label, FolderBrowserDialog, RadioButton
from System import Array


#This form just essentially lets the user designate stop and start frequency as well as step size
class SYZParams(Form):
    def __init__(self):
        self.Text = "Set SYZ Params"
        self.Width = 400
        self.Height = 300 

        self.startLabel = Label()
        self.startLabel.Text = "Start Frequency (GHz)"
        self.startLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.startLabel.Top = 10
        self.startLabel.Left = 50

        self.startTextbox = TextBox()
        self.startTextbox.Top = 10
        self.startTextbox.Left = self.startLabel.Right + 10

        self.stopLabel = Label()
        self.stopLabel.Text = "Stop Frequency (GHz)"
        self.stopLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.stopLabel.Top = self.startLabel.Bottom + 10
        self.stopLabel.Left = 50

        self.stopTextbox = TextBox()
        self.stopTextbox.Top = self.startLabel.Bottom + 10
        self.stopTextbox.Left = self.stopLabel.Right + 10

        self.stepLabel = Label()
        self.stepLabel.Text = "Step Size"
        self.stepLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.stepLabel.Top = self.stopLabel.Bottom + 10
        self.stepLabel.Left = 50

        self.stepTextbox = TextBox()
        self.stepTextbox.Top = self.stopLabel.Bottom + 10
        self.stepTextbox.Left = self.stepLabel.Right + 10

        self.setParams = Button()
        self.setParams.Text = "Set Parameters"
        self.setParams.Width = 300
        self.setParams.Top = self.stepLabel.Bottom + 10
        self.setParams.Left = 50
        self.setParams.Click += self.params


        self.closeButton = Button()
        self.closeButton.Text = "Next Step"
        self.closeButton.Width = 300
        self.closeButton.Top = self.setParams.Bottom + 10
        self.closeButton.Left = 50
        self.closeButton.Click += self.close

        self.Controls.Add(self.startLabel)
        self.Controls.Add(self.startTextbox)
        self.Controls.Add(self.stopLabel)
        self.Controls.Add(self.stopTextbox)
        self.Controls.Add(self.stepLabel)
        self.Controls.Add(self.stepTextbox)
        self.Controls.Add(self.setParams)
        self.Controls.Add(self.closeButton)

    def close(self,sender,e):
        self.Close()

    def params(self,sender,e):
        try:
            self.start = float(self.startTextbox.Text)
            self.stop = float(self.stopTextbox.Text)
            self.step = int(self.stepTextbox.Text)
        except:
            oDoc.ScrLogMessage("ERROR: PLEASE INPUT VALID NUMBERS")
            sys.exit()


    def returnParams(self):
        return [self.start,self.stop,self.step]

#Get the diestack information and change the names
class DieStack(Form):
    def __init__(self):
        self.Text = "Edit DieStack Name"
        self.Width = 400
        self.Height = 300

        self.dieBox = ListBox()
        self.dieBox.SelectionMode= SelectionMode.One
        self.dieBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.dieBox.Height = 80
        self.dieBox.Width = 300
        self.dieBox.Top = 10
        self.dieBox.Left = 50
        self.dieBox.HorizontalScrollbar = True


        partNames = set()
        dieStack = set()
        for components in oDoc.ScrGetComponentList("discrete devices"):
            part,refDes = components.split()
            if part not in partNames:
                partNames.add(part)
            if refDes not in dieStack:
                dieStack.add(refDes)
                self.dieBox.Items.Add(refDes)

        self.dieNameLabel = Label()
        self.dieNameLabel.Text = "Identify First APU then DIMM"
        self.dieNameLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.dieNameLabel.Top = self.dieBox.Bottom + 10
        self.dieNameLabel.Left = 50

        self.newDie = TextBox()
        self.newDie.Top = self.dieBox.Bottom + 10
        self.newDie.Left = self.dieNameLabel.Right + 10

        self.appendKey = Button()
        self.appendKey.Text = "Append Die Change"
        self.appendKey.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.appendKey.Top = self.newDie.Bottom + 10
        self.appendKey.Left = 50
        self.appendKey.Width = 300
        self.appendKey.Click += self.appendDie

        self.appendList = ListBox()
        self.appendList.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.appendList.Height = 50
        self.appendList.Width = 300
        self.appendList.Top = self.appendKey.Bottom + 10
        self.appendList.Left = 50
        self.appendList.HorizontalScrollbar = True


        self.button = Button()
        self.button.Text = "Change Die Names"
        self.button.Click += self.changeDie
        self.button.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.button.Top = self.appendList.Bottom + 10
        self.button.Width = 300   # Adjusting width to fit within form
        self.button.Left = 50

        self.closeButton = Button()
        self.closeButton.Text = "Close"
        self.closeButton.Width = 300
        self.closeButton.Top = self.button.Bottom + 10
        self.closeButton.Left = 50
        self.closeButton.Click += self.close
        self.Controls.Add(self.closeButton)

        self.Controls.Add(self.dieBox)
        self.Controls.Add(self.dieNameLabel)
        self.Controls.Add(self.newDie)
        self.Controls.Add(self.appendKey)
        self.Controls.Add(self.appendList)
        self.Controls.Add(self.button)
        self.Controls.Add(self.closeButton)

        self.dieDict = dict()

    def appendDie(self,sender,e):

        #Essentially starts to list out Die Name information
        replace_word = self.dieBox.SelectedItem
        keyWord = self.newDie.Text.strip()
        self.newDie.Clear()
        self.dieBox.ClearSelected()

        #Adds each user input into the dictionary
        if keyWord != "APU":
            val = int(replace_word[1:])
            for x in range(0,20):
                replace_word = replace_word[:1] + str(val)
                if replace_word not in self.dieDict.keys():
                    self.dieDict[replace_word] = keyWord + str(x)
                    transform = replace_word + " to " + keyWord + str(x)
                    self.appendList.Items.Add(transform)
                val = val + 1
        else:
            self.dieDict[replace_word] = keyWord
            transform = replace_word + " to " + keyWord
            self.appendList.Items.Add(transform)
        

    def close(self,sender,e):
        self.Close()

    def changeDie(self,sender,event):

        oDoc.Save()
        filePath = oDoc.GetFilePath()
        oDoc.ScrCloseProjectNoSave()

        #This opens the .siw file manually and changes the die stack through the text file
        for replace_word in self.dieDict.keys():
            with open(filePath, 'r') as input_file:
                lines = input_file.readlines()
    
            # Modify the content in memory
            modified_lines = []
            for l in lines:
                if "SET_SYZ_FWS_PARAMS" in l or "SIWAVE_OPTS_USE_CUSTOM_PI_SI" in l  or "SIWAVE_SYZ_EXPORT_TOUCHSTONE_FILE" in l:
                    newLine = l.replace("0","1")
                    modified_lines.append(newLine)
                else:
                    newLine = re.sub(r'(\s|^|")' + re.escape(replace_word) + r'(\s|$|")', r'\1' + self.dieDict[replace_word] + r'\2', l)
                    modified_lines.append(newLine)


            # Overwrite the input file with the modified content
            with open(filePath, 'w') as output_file:
                output_file.writelines(modified_lines)

        #Open the project once again
        oApp.OpenProject(filePath)

        self.dieBox.Items.Clear()
        
        partNames = set()
        dieStack = set()

        #Get the new die stack names
        for components in oDoc.ScrGetComponentList("discrete devices"):
            part,refDes = components.split()
            if part not in partNames:
                partNames.add(part)
            if refDes not in dieStack:
                dieStack.add(refDes)
                self.dieBox.Items.Add(refDes)
    
    def returnDieDict(self):
        return self.dieDict





class EditForm(Form):
    def __init__(self):
        self.Text = "Edit Layers and Materials"
        self.Width = 400
        self.Height = 300

        self.stkBox = ListBox()
        self.stkBox.SelectionMode= SelectionMode.One
        self.stkBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.stkBox.Height = 20
        self.stkBox.Width = 300
        self.stkBox.Top = 10
        self.stkBox.Left = 50
        self.stkBox.HorizontalScrollbar = True

        self.button = Button()
        self.button.Text = "Export STK File"
        self.button.Click += self.button_click_save
        self.button.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.button.Top = self.stkBox.Height + 10
        self.button.Width = 300   # Adjusting width to fit within form
        self.button.Left = 50

        self.layersBox = ListBox()
        self.layersBox.SelectionMode= SelectionMode.MultiExtended
        self.layersBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.layersBox.Height = 100
        self.layersBox.Width = 300
        self.layersBox.Top = 10 + self.button.Bottom
        self.layersBox.Left = 50
        self.layersBox.HorizontalScrollbar = True

        self.lowLoss = RadioButton()
        self.lowLoss.Text = "Low Loss"
        self.lowLoss.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.lowLoss.Top = 10 + self.layersBox.Bottom
        self.lowLoss.Left = 50

        self.mediumLoss = RadioButton()
        self.mediumLoss.Text = "Medium Loss"
        self.mediumLoss.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.mediumLoss.Top = 10 + self.layersBox.Bottom
        self.mediumLoss.Left = self.lowLoss.Right + 10

        self.highLoss = RadioButton()
        self.highLoss.Text = "High Loss"
        self.highLoss.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.highLoss.Top = 10 + self.layersBox.Bottom
        self.highLoss.Left = self.mediumLoss.Right + 10

        self.rough_button = Button()
        self.rough_button.Text = "Set Huray Roughness for Selected Layers"
        self.rough_button.Click += self.roughness
        self.rough_button.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.rough_button.Top = self.lowLoss.Bottom + 10
        self.rough_button.Width = 300   # Adjusting width to fit within form
        self.rough_button.Left = 50

        self.materialsBox = ListBox()
        self.materialsBox.SelectionMode= SelectionMode.One
        self.materialsBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.materialsBox.Height = 100
        self.materialsBox.Width = 300
        self.materialsBox.Top = 10 + self.rough_button.Bottom
        self.materialsBox.Left = 50
        self.materialsBox.HorizontalScrollbar = True

        self.name_label = Label()
        self.name_label.Text = "Name:"
        self.name_label.Top = self.materialsBox.Bottom + 10
        self.name_label.Left = 50

        self.name_text = TextBox()
        self.name_text.Top = self.name_label.Top
        self.name_text.Left = self.name_label.Right + 10

        self.frequency_label = Label()
        self.frequency_label.Text = "Material Frequency (GHz):"
        self.frequency_label.Top = self.name_label.Bottom + 10
        self.frequency_label.Left = 50

        self.frequency_text = TextBox()
        self.frequency_text.Top = self.frequency_label.Top
        self.frequency_text.Left = self.frequency_label.Right + 10

        self.permittivity_label = Label()
        self.permittivity_label.Text = "Typical Permittivity:"
        self.permittivity_label.Top = self.frequency_label.Bottom + 10
        self.permittivity_label.Left = 50

        self.permittivity_textbox = TextBox()
        self.permittivity_textbox.Top = self.permittivity_label.Top
        self.permittivity_textbox.Left = self.permittivity_label.Right + 10

        self.loss_tangent_label = Label()
        self.loss_tangent_label.Text = "Loss Tangent:"
        self.loss_tangent_label.Top = self.permittivity_label.Bottom + 10
        self.loss_tangent_label.Left = 50

        self.loss_tangent_textbox = TextBox()
        self.loss_tangent_textbox.Top = self.loss_tangent_label.Top
        self.loss_tangent_textbox.Left = self.loss_tangent_label.Right + 10

        #Add Conductor Box and the Conductor Materials in the box 
        self.conBox = ListBox()
        self.conBox.SelectionMode = SelectionMode.One
        self.conBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.conBox.Height = 100
        self.conBox.Width = 300
        self.conBox.Top = 10 + self.loss_tangent_label.Bottom
        self.conBox.Left = 50
        self.conBox.HorizontalScrollbar = True

        self.clone = Button()
        self.clone.Text = "Add Conductor Material and New Material"
        self.clone.Click += self.clone_material_button
        self.clone.Top = self.conBox.Bottom + 10
        self.clone.Left = 50
        self.clone.Width = 300

        self.closeButton = Button()
        self.closeButton.Text = "Next Step"
        self.closeButton.Width = 300
        self.closeButton.Top = self.clone.Bottom + 10
        self.closeButton.Left = 50
        self.closeButton.Click += self.close

        self.Controls.Add(self.stkBox)
        self.Controls.Add(self.button)
        self.Controls.Add(self.layersBox)
        self.Controls.Add(self.lowLoss)
        self.Controls.Add(self.mediumLoss)
        self.Controls.Add(self.highLoss)
        self.Controls.Add(self.rough_button)
        self.Controls.Add(self.materialsBox)
        self.Controls.Add(self.name_label)
        self.Controls.Add(self.name_text)
        self.Controls.Add(self.frequency_label)
        self.Controls.Add(self.frequency_text)
        self.Controls.Add(self.permittivity_label)
        self.Controls.Add(self.permittivity_textbox)
        self.Controls.Add(self.loss_tangent_label)
        self.Controls.Add(self.loss_tangent_textbox)
        self.Controls.Add(self.conBox)
        self.Controls.Add(self.clone)
        self.Controls.Add(self.closeButton)

    def close(self,sender,e):
        self.Close()

    def button_click_save (self, sender, e):

        #Export the file
        file_dialog = SaveFileDialog()
        file_dialog.Filter = "Stackup files (*.stk)|*.stk|All files (*.*)|*.*"  # Set filter for .stk files
        result = file_dialog.ShowDialog()

        if result == DialogResult.OK:
            self.selected_file = file_dialog.FileName
            oDoc.ScrExportLayerStackup(self.selected_file)
            self.stkBox.Items.Clear()
            self.stkBox.Items.Add(self.selected_file)
            oDoc.ScrImportLayerStackup(self.selected_file)
        
        self.layersBox.Items.Clear()

        #Get all the current material already
        for x in oDoc.ScrGetLayerNameList():
            material = oDoc.ScrGetLayerMaterial(x)
            self.layersBox.Items.Add(x)
            self.materialsBox.Items.Add(material)

        #Get all the conductor materials
        for x in conductors:
            self.conBox.Items.Add(x)

        self.files = [self.selected_file,self.selected_file,self.selected_file]
        
        
    def clone_material_button (self,sender,e):
        material = self.materialsBox.SelectedItem
        conduct = self.conBox.SelectedItem
        name = self.name_text.Text
        freq = self.frequency_text.Text
        typical = self.permittivity_textbox.Text
        loss_tangent = self.loss_tangent_textbox.Text
        filename = str(self.selected_file)

        try:
            float(typical)
            float(loss_tangent)
            freq = int(freq) * 1000000000
        except:
            oDoc.ScrLogMessage("ERROR: PLEASE INPUT VALID NUMBERS")
            sys.exit()

        #Open the stackup file and modify the contents based off user input
        with open(filename, 'r') as file:
            stk_content = file.read()

        newMaterial = "\t\t$begin 'Insulator'\n\t\t\tName='{}'\n\t\t\tPermittivity={}\n\t\t\tLossTangent={}\n\t\t\tDerivedFrom='vacuum'\n\t\t\tAnsoftID=0\n\t\t\t$begin 'Djordjevic-Sarkar'\n\t\t\t\tMeasurementFreq={}\n\t\t\t\tSigmaDC=1e-12\n\t\t\t\tEpsDC=0\n\t\t\t$end 'Djordjevic-Sarkar'\n\t\t$end 'Insulator'\n".format(name,typical,loss_tangent,freq)
        startIndex = stk_content.find("$begin 'Materials'")
        insertion = startIndex + len("$begin 'Materials'") + 1

        newContent = stk_content[:insertion] + newMaterial + stk_content[insertion:]

        with open(filename,'w') as file:
            file.write(newContent)
        
        oDoc.ScrImportLayerStackup(filename)
        #Fix that it only edits the layers between top and bottom
        topLayer = False
        bottomLayer = False

        #Sets the layers and its material based on its type
        for layer in oDoc.ScrGetLayerNameList():
            if layer == "BOTTOM":
                oDoc.ScrSetLayerMaterial(layer,conduct)
                bottomLayer = True
            if topLayer:
                if topLayer and not bottomLayer:
                    if oDoc.ScrGetLayerType(layer) == 1: #Metal Type
                        oDoc.ScrSetMetalLayerFillerMaterial(layer,name)
                        oDoc.ScrSetLayerMaterial(layer,conduct) #conductor 
                    elif oDoc.ScrGetLayerType(layer) == 0: #Dielectric Type
                        oDoc.ScrSetLayerMaterial(layer,name)
            if layer == "TOP":
                oDoc.ScrSetLayerMaterial(layer,conduct)
                topLayer = True
            

            
            oDoc.ScrExportLayerStackup(filename)


    def roughness (self,sender,e):
        selected_layers = [item for item in self.layersBox.SelectedItems]
        filename = self.selected_file

        #Checks what the user inputted for the Huray Roughness MOdel
        if self.lowLoss.Checked:
            huray = self.lowLoss.Text
        elif self.mediumLoss.Checked:
            huray = self.mediumLoss.Text
        elif self.highLoss.Checked:
            huray = self.highLoss.Text

        with open(filename, 'r') as file:
            stk_content = file.read()

        # Split the content into individual lines
        lines = stk_content.split('\n')

        # Iterate through the lines to find and modify the layers as required
        for i in range(0,len(lines)):
            if 'LayerID' in lines[i] and 'LayerName' in lines[i]:
                layer_name = lines[i].split("LayerName='")[1].split("'")[0]
                if layer_name in selected_layers:
                    # Update TopRoughnessHurayModel and IsTopRoughnessHuray
                    lines[i] = re.sub(r"TopRoughnessHurayModel='[^']*'", "TopRoughnessHurayModel='{}'".format(huray), lines[i])
                    lines[i] = lines[i].replace("IsTopRoughnessHuray=false", "IsTopRoughnessHuray=true")

                    #Update BottomRoughnessHurayModel
                    lines[i] = lines[i] = re.sub(r"BottomRoughnessHurayModel='[^']*'", "BottomRoughnessHurayModel='{}'".format(huray), lines[i])
                    lines[i] = lines[i].replace("IsBottomRoughnessHuray=false", "IsBottomRoughnessHuray=true")

        # Join the modified lines back into the content
        modified_stk_content = '\n'.join(lines)

        # Write the modified content back to the .stk file
        with open(filename, 'w') as file:
            file.write(modified_stk_content)

        oDoc.ScrImportLayerStackup(filename)
    
    def returnFileNames(self):

        return self.selected_file



class ClippingForm(Form):
    def __init__(self):
        self.Text = "Clipping Form"
        self.Width = 400
        self.Height = 300

        self.my_dictionary = dict()

        self.search_box = TextBox()  # Create a search box
        self.search_box.Top = 10
        self.search_box.Left = 10
        self.search_box.Width = 200
        self.search_box.TextChanged += self.search_box_text_changed
        self.Controls.Add(self.search_box)

        self.netsBox = ListBox()
        self.netsBox.SelectionMode = SelectionMode.MultiExtended
        self.netsBox.Height = 200
        self.netsBox.Width = 200
        self.netsBox.Top = self.search_box.Bottom + 10
        self.netsBox.Left = 10
        self.netsBox.HorizontalScrollbar = True

        # Adding key box
        self.key_box = TextBox()
        self.key_box.Top = self.netsBox.Bottom + 10
        self.key_box.Left = 10
        self.key_box.Width = 200
        self.Controls.Add(self.key_box)

        for item in nets:
            self.netsBox.Items.Add(item)

        self.netsBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.Controls.Add(self.netsBox)

        self.button = Button()
        self.button.Width = 200
        self.button.Text = "Append Selected Nets"
        self.button.Click += self.button_click
        self.button.Top = self.key_box.Bottom + 10
        self.button.Left = 10
        self.button.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.Controls.Add(self.button)

        self.keylist = ListBox()
        self.keylist.Top = self.button.Bottom+10
        self.keylist.Left=10
        self.keylist.Width = 200
        self.keylist.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.keylist.HorizontalScrollbar = True
        self.Controls.Add(self.keylist)

        self.finishedKeyList = ListBox()
        self.finishedKeyList.Top = self.keylist.Bottom + 10
        self.finishedKeyList.Left = 10
        self.finishedKeyList.Width = 200
        self.finishedKeyList.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.finishedKeyList.HorizontalScrollbar = True
        self.Controls.Add(self.finishedKeyList)

        self.generate = Button()
        self.generate.Width = 200
        self.generate.Text = "Generate Files"
        self.generate.Click += self.generate_click
        self.generate.Top = self.finishedKeyList.Bottom + 10
        self.generate.Left = 10
        self.generate.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.Controls.Add(self.generate)

        self.closeButton = Button()
        self.closeButton.Text = "Close"
        self.closeButton.Width = 200
        self.closeButton.Top = self.generate.Bottom + 10
        self.closeButton.Left = 10
        self.closeButton.Click += self.close
        self.Controls.Add(self.closeButton)

        self.layersEdit = False
        self.typical = ""

        self.SYZParamsEdit = False
        self.start = 0
        self.stop = 0
        self.step = 0

        self.dieEdit = False
        self.dieDict = {}

    #This function lets the user search certain nets
    def search_box_text_changed(self, sender, e):
        search_text = self.search_box.Text.lower()
        self.netsBox.Items.Clear()
        for item in nets:
            if search_text in item.lower():
                self.netsBox.Items.Add(item)

    def button_click(self, sender, event):
        selected_items = [self.netsBox.Items[index] for index in self.netsBox.SelectedIndices]
        key = self.key_box.Text.strip()
        if key:
            self.my_dictionary[key] = selected_items
            oDoc.ScrLogMessage(key)
            oDoc.ScrLogMessage(str(selected_items))
        self.netsBox.ClearSelected()
        self.key_box.Clear()
        self.keylist.Items.Add(key)

    def generate_click(self, sender, event):

        #Start to develop each file 
        fileList = sorted(self.my_dictionary.keys())
        oDoc.ScrLogMessage(fileList)
        for x in fileList:
            curFile = fileName.split("\\")[-1]
            curFile = curFile.replace(".siw","")
            newFilename = r"\{}_pcb_1dimm_ch-{}_typZ.siw".format(curFile,x) #Make the new name of the file being generated
            oDoc.ScrLogMessage(newFilename)
            destination= destFolder + newFilename
            shutil.copy(fileName,destination)
            oApp.OpenProject(destination)

            if not self.layersEdit:
                editForm = EditForm()
                editForm.ShowDialog()
                self.typical = editForm.returnFileNames()
                self.layersEdit = True


            oDoc.ScrLogMessage("SCRIPT: Enabling Layer Visibility")

            for layer in oDoc.ScrGetLayerNameList():
                oDoc.ScrSetLayerVisibility(layer,1,1,1,1,1)
        
            oDoc.ScrLogMessage("SCRIPT: Finished Enabling Layer Visibility")
            oDoc.ScrLogMessage("SCRIPT: Number of selected nets:" + str(len(self.my_dictionary[x])))
            oDoc.ScrLogMessage("SCRIPT: Clipping Around Selected Nets")

            oDoc.ScrClipDesignAroundNets(self.my_dictionary[x],"1mm",True,0,True,False)

            oDoc.ScrLogMessage("SCRIPT: Finished Clipping Around Selected Nets")

            oDoc.ScrLogMessage("SCRIPT: Deleting UnWanted Nets")

            powerNets = oDoc.ScrGetPwrGndNetNameList()
            deleteNets = []
            for net in oDoc.ScrGetNetNameList():
                if net not in powerNets and net not in self.my_dictionary[x]:
                    deleteNets.append(net)
                
            oDoc.ScrDeleteNets(deleteNets)
            oDoc.ScrLogMessage("SCRIPT: Finished Deleting UnWanted Nets")

            oDoc.ScrLogMessage("SCRIPT: Selecting Nets")
            for y in self.my_dictionary[x]:
                oDoc.ScrSelectNet(y,1)

            oDoc.ScrLogMessage("SCRIPT: Finished Selecting Nets")

            oDoc.ScrImportLayerStackup(self.typical)
            oDoc.Save()

            
            #Change Die Names
            oDoc.ScrLogMessage("SCRIPT: Opening Die Form")

            if not self.dieEdit:
                dieForm = DieStack()
                dieForm.ShowDialog()
                self.dieDict = dieForm.returnDieDict()
                self.dieEdit = True

            oDoc.Save()
            filePath = oDoc.GetFilePath()
            oDoc.ScrCloseProjectNoSave()

            #Change the die names of the file
            for replace_word in self.dieDict.keys():
                with open(filePath, 'r') as input_file:
                    lines = input_file.readlines()
        
                # Modify the content in memory
                modified_lines = []
                for l in lines:
                    if "SET_SYZ_FWS_PARAMS" in l or "SIWAVE_OPTS_USE_CUSTOM_PI_SI" in l  or "SIWAVE_SYZ_EXPORT_TOUCHSTONE_FILE" in l:
                        newLine = l.replace("0","1")
                        modified_lines.append(newLine)
                    else:
                        newLine = re.sub(r'(\s|^|")' + re.escape(replace_word) + r'(\s|$|")', r'\1' + self.dieDict[replace_word] + r'\2', l)
                        modified_lines.append(newLine)

                # Overwrite the input file with the modified content
                with open(filePath, 'w') as output_file:
                    output_file.writelines(modified_lines)
            
            oDoc.ScrLogMessage("SCRIPT: Finished Die Name Change")

            oApp.OpenProject(filePath)

            oDoc.Save()

            oDoc.ScrLogMessage("SCRIPT: Creating Pin Groups")
            partNames=set()
            for component in oDoc.ScrGetComponentList("discrete devices"):
                part, refDes =component.split()
                if part not in partNames:
                    partNames.add(part)
            
            partNames = list(partNames)

            for p in partNames:
                oDoc.ScrCreatePinGroupByNet(p,"","GND",p+"_PinGroup",True)

            oDoc.ScrLogMessage("SCRIPT: Finished Creating Pin Groups")

            oDoc.ScrLogMessage("SCRIPT: Generating Ports on Pin Groups")
            oDoc.ScrPlacePortsAtPinsOnSelectedNets(50.0,"GND",True,[])
            oDoc.ScrLogMessage("SCRIPT: Finished Generating Ports on Pin Groups")

            oDoc.ScrLogMessage("SCRIPT: Configuring SYZ Parameters Simulation")
            if not self.SYZParamsEdit:
                syzForm = SYZParams()
                syzForm.ShowDialog()
                self.start, self.stop, self.step = syzForm.returnParams()
                self.SYZParamsEdit = True

            #Initiates the syz parameters such as stop and start frequencies as well as step size
            oDoc.ScrClearAllSweeps("syz")
            oDoc.ScrAppendSweep("syz",int(self.start*1000000000),int(self.stop*1000000000),self.step,False)
            oDoc.ScrSIwaveEnable_3D_DDM(0)
            oDoc.ScrEnableErcSimSetup(1)
            oDoc.ScrSetNumCpusToUse(16)


            #Sets area parameters 
            oDoc.ScrSetMinCutoutArea(10000,"mm2")
            oDoc.ScrSetMinPadAreaToMesh("0.1mm2")
            oDoc.ScrSetMinPlaneAreaToMesh("0.01mm2")

            oDoc.Save()

            oDoc.ScrLogMessage("SCRIPT: Finished Configuring SYZ Parameters Simulation")

            # oDoc.ScrRunValidationCheck()
            # oDoc.ScrRunSimulation("syz", "SYZ SWEEP")
            oDoc.Save()

            self.keylist.Items.Remove(x)
            self.finishedKeyList.Items.Add(x)
            del self.my_dictionary[x]

            
            
    def close(self,sender,e):
        self.Close()


class FileForm(Form):
    def __init__(self):
        self.Text = "Open SIW File & Set Destination"
        self.Width = 400
        self.Height = 300
        
        self.siwFileBox = ListBox()
        self.siwFileBox.Height = 25
        self.siwFileBox.Width = 300
        self.siwFileBox.Top = 10
        self.siwFileBox.Left = 50
        self.siwFileBox.HorizontalScrollbar = True
        self.siwFileBox.SelectionMode= SelectionMode.One

        self.siwBrowse = Button()
        self.siwBrowse.Width = 300
        self.siwBrowse.Text = "Browse SIW File"
        self.siwBrowse.Left = 50
        self.siwBrowse.Top = self.siwFileBox.Bottom + 10
        self.siwBrowse.Click += self.selectSIW

        self.destFolder = ListBox()
        self.destFolder.Height = 25
        self.destFolder.Width = 300
        self.destFolder.Top = self.siwBrowse.Bottom + 10
        self.destFolder.Left = 50
        self.destFolder.HorizontalScrollbar = True
        self.destFolder.SelectionMode= SelectionMode.One

        self.destBrowse = Button()
        self.destBrowse.Text = "Destination Folder"
        self.destBrowse.Width = 300
        self.destBrowse.Top = self.destFolder.Bottom + 10
        self.destBrowse.Left = 50
        self.destBrowse.Click += self.selectFolder

        self.closeButton = Button()
        self.closeButton.Text = "Next Step"
        self.closeButton.Width = 300
        self.closeButton.Top = self.destBrowse.Bottom + 10
        self.closeButton.Left = 50
        self.closeButton.Click += self.close

        self.Controls.Add(self.siwFileBox)
        self.Controls.Add(self.siwBrowse)
        self.Controls.Add(self.destFolder)
        self.Controls.Add(self.destBrowse)
        self.Controls.Add(self.closeButton)

    def selectSIW (self, sender, e):
        siwFile = OpenFileDialog()
        result = siwFile.ShowDialog()
        
        if result == DialogResult.OK:
            self.file = siwFile.FileName
            self.siwFileBox.Items.Clear()
            self.siwFileBox.Items.Add(self.file)

    def selectFolder (self,sender,e):
        folder = FolderBrowserDialog()
        result = folder.ShowDialog()

        if result == DialogResult.OK:
            self.selectedFolder = folder.SelectedPath
            self.destFolder.Items.Clear()
            self.destFolder.Items.Add(self.selectedFolder)
    
    def close(self,sender,e):
        self.Close()


    def getLocations(self):
        return self.file, self.selectedFolder

if __name__ == "__main__":
    StartForm = FileForm()
    StartForm.ShowDialog()


    fileName, destFolder = StartForm.getLocations()

    #After getting the file, essentially list out all the materials for further use
    with open(fileName, 'r') as file:
        netCollecting = False
        conCollecting = False
        firstQuote = False
        conductors = []
        nets = []
        for line in file:
            if line.strip() == "B_MATERIALS":
                conCollecting = True
            elif line.strip() == "E_MATERIALS":
                conCollecting = False
            
            if conCollecting:
                if "CONDUCTOR" in line.strip() and "#" not in line.strip():
                    words = line.strip().split('"')
                    conductorMat = words[1]
                    conductors.append(conductorMat)

            if line.strip() == "B_NETS":
                netCollecting = True
            elif line.strip() == "E_NETS":
                netCollecting = False
                break  # Stop collecting lines once we encounter "E_NETS"
            elif netCollecting:
                words = line.strip().split(" ")
                nets.append(words[1])


    form = ClippingForm()
    form.ShowDialog()


