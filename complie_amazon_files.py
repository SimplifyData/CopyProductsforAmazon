import pandas as pd


def compile_amz_file():

name_output_file = input("Name the output file: ")

file = active_wkbk()

try:

    file_num_max = input(" Total number of files to be combined?: ")

    file_num_max = int(file_num_max)

except:

    print("Please enter a whole number (integer)")



for num in range(1,file_num_max+1):
    if num <= 1:
        active_wkbk(str(file.split()[0] + " (" + str(num) + ")." + str(file.split(".")[1])))

        data  = Cell(all_sheets()[0],(1,1)).table

        df = pd.DataFrame(data[1:], columns= data[0])

        valid_data_df = df.ix[:,:][ df.ix[:,1] == 0]

        new_wkbk()

        active_wkbk(all_wkbks()[1])

        CellRange(all_sheets()[0],(1,1),(1,88)).value = data[0]

        Cell(all_sheets()[0],(3,1)).table = valid_data_df.values.tolist()

        save(str(path)+ "\\" + str(name_output_file) )

    if num >1:
        open_wkbk(

            str(file.split()[0] + " (" + str(num) + ")." + str(file.split(".")[1]))
        )

