import os
import shutil
import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application, Form, Button, ListBox, SelectionMode, AnchorStyles, TextBox, OpenFileDialog, DialogResult, SaveFileDialog, Label, FolderBrowserDialog
from System import Array
import sys

class highLowForm(Form):
    def __init__(self):
        self.Text = "Corner Modeling"
        self.Width = 400
        self.Height = 300

        self.folderBox = ListBox()
        self.folderBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.folderBox.Height = 20
        self.folderBox.Width = 350
        self.folderBox.Top = 10
        self.folderBox.Left = 50

        self.folderButton = Button()
        self.folderButton.Text = "Select Folder" 
        self.folderButton.Width = 350
        self.folderButton.Height = 20
        self.folderButton.Top = self.folderBox.Bottom + 10
        self.folderButton.Left = 50
        self.folderButton.Click += self.folderClick

        self.layerBox = ListBox()
        self.layerBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.layerBox.Height = 20
        self.layerBox.Width = 350
        self.layerBox.Top = self.folderButton.Bottom + 40
        self.layerBox.Left = 50

        self.layerButton = Button()
        self.layerButton.Text = "Select Stackup File"
        self.layerButton.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.layerButton.Width = 350
        self.layerButton.Height = 20
        self.layerButton.Top = self.layerBox.Bottom + 10
        self.layerButton.Left = 50
        self.layerButton.Click += self.layerClick

        self.materials = ListBox()
        self.materials.SelectionMode = SelectionMode.One
        self.materials.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.materials.Height = 50
        self.materials.Width = 350
        self.materials.Top = self.layerButton.Bottom + 10
        self.materials.Left = 50


        self.high_label = Label()
        self.high_label.Text = "High Permittivity:"
        self.high_label.Top = self.materials.Bottom + 10
        self.high_label.Left = 50

        self.high_text = TextBox()
        self.high_text.Top = self.materials.Bottom + 10
        self.high_text.Left = self.high_label.Right + 10

        self.high_tan_label = Label()
        self.high_tan_label.Text = "High Loss Tangent:"
        self.high_tan_label.Top = self.materials.Bottom + 10
        self.high_tan_label.Left = self.high_text.Right + 10

        self.high_tan_text = TextBox()
        self.high_tan_text.Top = self.materials.Bottom + 10
        self.high_tan_text.Left = self.high_tan_label.Right + 10 

        self.low_label = Label()
        self.low_label.Text = "Low Permittivity:"
        self.low_label.Top = self.high_label.Bottom + 10
        self.low_label.Left = 50

        self.low_text = TextBox()
        self.low_text.Top = self.high_label.Bottom + 10
        self.low_text.Left = self.low_label.Right + 10

        self.low_tan_label = Label()
        self.low_tan_label.Text = "Low Loss Tangent:"
        self.low_tan_label.Top = self.high_label.Bottom + 10
        self.low_tan_label.Left = self.low_text.Right + 10
        
        self.low_tan_text = TextBox()
        self.low_tan_text.Top = self.high_label.Bottom + 10
        self.low_tan_text.Left = self.low_tan_label.Right + 10

        self.generateButton = Button()
        self.generateButton.Text = "Generate High and Low Files"
        self.generateButton.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.generateButton.Width = 350
        self.generateButton.Height = 30
        self.generateButton.Top = self.low_label.Bottom + 10
        self.generateButton.Left = 50
        self.generateButton.Click += self.generateFile

        self.Controls.Add(self.folderBox)
        self.Controls.Add(self.folderButton)
        self.Controls.Add(self.layerBox)
        self.Controls.Add(self.layerButton)
        self.Controls.Add(self.materials)
        self.Controls.Add(self.high_label)
        self.Controls.Add(self.high_text)
        self.Controls.Add(self.high_tan_label)
        self.Controls.Add(self.high_tan_text)
        self.Controls.Add(self.low_tan_label)
        self.Controls.Add(self.low_tan_text)
        self.Controls.Add(self.low_label)
        self.Controls.Add(self.low_text)
        self.Controls.Add(self.generateButton)

    def folderClick(self,sender,e):
            folder = FolderBrowserDialog()
            result = folder.ShowDialog()

            if result == DialogResult.OK:
                self.selectedFolder = folder.SelectedPath
                self.folderBox.Items.Clear()
                self.folderBox.Items.Add(self.selectedFolder)
    
    def layerClick(self,sender,e):
         layerFile = OpenFileDialog()
         result = layerFile.ShowDialog(None)
         
         if result == DialogResult.OK:
              self.selectedLayer = layerFile.FileName
              self.layerBox.Items.Clear()
              self.layerBox.Items.Add(self.selectedLayer)

              with open(self.selectedLayer,'r') as f:
                   stk_content = f.read()

              lines = stk_content.split("\n")
              
              for l in lines:
                if "\tName=" in l:
                     item = l.replace("\t","")
                     item = item.replace("Name=","")
                     self.materials.Items.Add(item)

              

    def generateFile(self,sender,e):
        #Generate the new high and low stk files
        highStk =  self.selectedLayer
        highStk = highStk.replace(".stk","high.stk")
        lowStk = self.selectedLayer
        lowStk = lowStk.replace(".stk","low.stk")
        selectedMaterial = self.materials.SelectedItem

        #Makes the copies of the two variants of the typical layer stackup file
        shutil.copy(self.selectedLayer,highStk)
        shutil.copy(self.selectedLayer,lowStk)

        for x in [lowStk,highStk]:
            if "high.stk" in x:
                permittivity = self.high_text.Text
                loss = self.high_tan_text.Text
            else:
                permittivity = self.low_text.Text
                loss = self.low_tan_text.Text
            
            try:
                float(permittivity)
                float(loss)
            except:
                oDoc.ScrLogMessage("ERROR: PLEASE INPUT VALID NUMBERS")
                sys.exit()

            #Edit the high stack and low stack file with the selected material
            with open(x, 'r') as f:
                stk_content = f.read()
            
            lines = stk_content.split('\n')
            foundMaterial = False
            modified = []

            for l in lines:
                if "\tName={}".format(selectedMaterial) in l:
                    foundMaterial = True

                if foundMaterial and "Permittivity=" in l:
                    modified.append("\t\t\tPermittivity={}\n".format(permittivity))
                elif foundMaterial and "LossTangent=" in l:
                    modified.append("\t\t\tLossTangent={}\n".format(loss))
                    foundMaterial = False
                else:
                    modified.append(l+"\n")

            with open(x, 'w') as f:
                f.writelines(modified)

        

        #After generating the two stackup files, now actually copy the typical .siw file and import the correspanding .stk file
        for file in os.listdir(self.selectedFolder):
            if ".siw" in file:
                for fileType in ["lowZ","highZ"]:
                    newFile = str(file)
                    newFile = newFile.replace("typZ",fileType)
                    original = self.selectedFolder + r"\{}".format(file)
                    location = self.selectedFolder + r"\{}".format(newFile)
                    
                    shutil.copy(original,location)

                    replaceWord = file.replace(".siw","")
                    newWord = newFile.replace(".siw","")

                    if fileType == "lowZ":
                        importLayer = lowStk
                    else:
                        importLayer = highStk

                    oDoc.ScrLogMessage("Opening {}".format(fileType))          
                    oApp.OpenProject(location)
                    oDoc.ScrLogMessage("Importing Layer")
                    val = oDoc.ScrImportLayerStackup(importLayer)
                    if val == 0:
                        oDoc.ScrLogMessage("FAILED Import")
                    oDoc.Save()
                    oDoc.ScrCloseProject()

                    with open(location, 'r') as input_file:
                        lines = input_file.readlines()

                    # Modify the content in memory
                    modified_lines = []
                    for l in lines:
                        newLine = l.replace(replaceWord,newWord)
                        modified_lines.append(newLine)


                    # Overwrite the input file with the modified content
                    with open(location, 'w') as output_file:
                        output_file.writelines(modified_lines)






if __name__ == "__main__":
     cornerModels = highLowForm()
     cornerModels.ShowDialog()


