import tkinter as tk
from tkinter import filedialog
import os

def process_files():
    input_directory = input_entry.get()
    output_directory = output_entry.get()

    #After the inputs of both the input and output directory, get the touchstone files
    if input_directory and output_directory:

        #For each input, make a 
        for filename in os.listdir(input_directory):
            input_file_path = os.path.join(input_directory, filename)
            output_file_path = os.path.join(output_directory, "S2eye_" + filename)

            #open the input file and change the contents accordingly
            with open(input_file_path, 'r') as input_file:
                lines = input_file.readlines()

            modified = []
            if "dq0-31" in filename:
                current_DQ = 0
            else:
                current_DQ = 32
            for line in lines:
                if "Port" in line:
                    contents = line.split()
                    data = contents[3]
                    if "CHECK" in data:
                        contents[3] = createCheck(data)
                    elif "DATA" in data:
                        contents[3] = createData(data, current_DQ)
                        if "DIMM" in contents[3]:
                            current_DQ += 1
                    else:  # DQ
                        contents[3] = createDQ(data)
                    contents.append("\n")

                    modified.append(" ".join(contents))
                else:
                    modified.append(line)

            #Open the output file and input the modified contents
            with open(output_file_path, 'w') as output_file:
                output_file.writelines(modified)

def select_input_directory():
    input_directory = filedialog.askdirectory()
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_directory)

def select_output_directory():
    output_directory = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_directory)

def createCheck(checkLine):
    checkLine = checkLine.split("_")
    newLine = []
    newLine.append(checkLine[1][:2])
    newLine.append(checkLine[2])
    newLine.append(checkLine[3])
    newLine.append(checkLine[5])
    
    return "_".join(newLine)

def createData(dataLine, num):
    dataLine = dataLine.split("_")
    newLine = []
    newLine.append(dataLine[1][:2])
    newLine.append("DQ")
    newLine.append(str(num))
    newLine.append(dataLine[5])

    return "_".join(newLine)

def createDQ(DQline):
    DQline = DQline.split("_")
    newLine = []
    newLine.append(DQline[1][:2])
    newLine.append("DQS")
    newLine.append(DQline[3])
    newLine.append(DQline[2].replace("DQS", ""))
    newLine.append(DQline[4])
    return "_".join(newLine)


if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    root.title("File Processing")

    # Create input directory selection
    input_label = tk.Label(root, text="Input Directory:")
    input_label.grid(row=0, column=0, padx=10, pady=5)

    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5)

    input_button = tk.Button(root, text="Browse", command=select_input_directory)
    input_button.grid(row=0, column=2, padx=10, pady=5)

    # Create output directory selection
    output_label = tk.Label(root, text="Output Directory:")
    output_label.grid(row=1, column=0, padx=10, pady=5)

    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5)

    output_button = tk.Button(root, text="Browse", command=select_output_directory)
    output_button.grid(row=1, column=2, padx=10, pady=5)

    # Create process button
    process_button = tk.Button(root, text="Process Files", command=process_files)
    process_button.grid(row=2, column=1, padx=10, pady=20)

    # Run the application
    root.mainloop()
