from __future__ import print_function
import arcpy, os, numpy
import pandas as pd
arcpy.env.overwriteOutput=True

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "AHPython Toolbox"
        self.alias = "ahp"

        # List of tool classes associated with this toolbox
        self.tools = [AHP]


class AHP(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Analytical Hierarchy Process Calculation Tool"
        self.description = "This tool was developed to use the Analytical Hierarchy Process (ahp) method to calculate the weights of each factor compared to the other factors."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

## The First Parameter takes the factors 
        
        param0 = arcpy.Parameter()
        param0.name= "parameter0"
        param0.displayName="Factors Names"
        param0.direction="Input"
        param0.datatype="DEFile"
        param0.parameterType="Required"
        param0.multiValue=True

## The Second Parameter takes the user desired output location and output name
        
        param1 = arcpy.Parameter()
        param1.name= "parameter1"
        param1.displayName="Output File"
        param1.direction="Output"
        param1.datatype="DEFile"
        param1.parameterType="Required"

## The Third Parameter is a Boolean that takes the User preference regarding his desire to export his result to a Terrset compatible files that will accompany the output file

        param2 = arcpy.Parameter()
        param2.name = "parameter2"
        param2.displayName="Create a (.pcf) and (.dsf)  files"
        param2.direction="input"
        param2.datatype="GPBoolean"
        param2.parameterType="Optional"
 
        return [param0, param1,param2]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

## Check if the number of inserted factors is below or above the limits of the Tool
        
        if parameters[0].altered and parameters[1].hasBeenValidated:
            factorsinput=parameters[0].valueAsText
            factors=factorsinput.split(";")
            if len(factors)>=15:
                parameters[0].setErrorMessage("The chosen factors are above the limits of this tool, please reduce thier number.")
            elif len(factors) < 3:
                parameters[0].setErrorMessage("The chosen factors are below the limits of this tool, please increase thier number.")                                
        return
        
        

    def execute(self, parameters, messages):
        """The source code of the tool."""

## Reading Factors from File

        factorsinput=parameters[0].valueAsText
        factors=factorsinput.split(";")
        opath=parameters[1].valueAsText
        if opath[-4:]!=".xls":
            opath=opath+".xls"
        else:
            pass
        No_Factors = len(factors)
        x=0
        while x<No_Factors:
            factors[x]=os.path.splitext(os.path.basename(factors[x]))[0]
            x+=1
            
## Create the main table that will hold the factors and their weights based on the user input

        com_board = [["" for i in range(No_Factors)] for z in range(No_Factors)]
        n=0
        while n < No_Factors:
            com_board[n][n] = 1
            n += 1
        altered_Block=((No_Factors**2)-No_Factors)/2
        x = 0
        while x < altered_Block:
            z = x + 1
            while z < No_Factors:
                com_board[x][z] = "X"
                z += 1
            x += 1
        print_table = pd.DataFrame(data=com_board, index=factors, columns=factors, dtype=numpy.float64)

## Write the newly created table to a Microsoft Excel File

        writer = pd.ExcelWriter(opath)
        print_table.to_excel(writer,'Judgment Value Matrix')
        writer.save()
        

## Start the Microsoft Excel application to allow the user to insert the factors proirity in a pairwise comparision table
        
        os.system('start excel.exe {}'.format(opath))
        time.sleep(15)

## Check to see if the User still have the Excel file open or not
        
        x=0
        z=0
        while x<1:
            try:
                time.sleep(7)
                os.rename(opath,opath)
                x+=1
            except WindowsError:
                if z==2:
                    messages.addMessage("Excel still open, please enter the values (and the fraction should written in Decimals) in the empty space then save and close the file to continue")
                    z=0
                z+=1

## Once the program verified that the file is closed, it will read the file and load its content into the dataframe inside the program for further prossesing

        print_table = pd.read_excel(opath)
        x = 0
        while x < altered_Block:
            z = x + 1
            while z < No_Factors:
                print_table.iat[x,z] = 1.0/print_table.iat[z,x]
                z += 1
            x += 1

## then overwrite the processed content into the same Excel file and save it into the first sheet under the 'Judgment Value Matrix' name

        writer = pd.ExcelWriter(opath)
        print_table.to_excel(writer,'Judgment Value Matrix')


## Calculate Total value for each Column from the processed excel file
        x = 0
        tot_factors = [0 for i in range(No_Factors)]
        while x < No_Factors:
            j = 0
            while j < No_Factors:
                tot_factors[x] = tot_factors[x] + print_table.iat[j, x]
                j += 1
            x += 1

## Calculate the Normalized values for Each Factor and save it into the second sheet under the 'Normalized Values' name
            
        norm_factors = print_table.copy(deep=True).astype("float64", copy=True)
        x = 0
        while x < No_Factors:
            j = 0
            while j < No_Factors:
                norm_factors.iat[j, x ] = (print_table.iat[j, x]+0.00000000001) / tot_factors[x]
                j += 1
            x += 1
        norm_factors.to_excel(writer,'Normalized Values')

## Calculate the Priority vector or weight Values and save it into the third sheet under the 'Priority Vector or Weights' name
        
        priority_vector = norm_factors.sum(axis=1) / No_Factors
        p_vector=priority_vector.to_frame(name="Priority vector or Weight")
        p_vector.to_excel(writer, 'Priority Vector or Weights')


## Calculate Consistancy Ratio
        x = 0
        Eigen_Value = 0
        factors_temp=[i for i in range(No_Factors)]
        z=0
        while z<No_Factors:
            factors_temp[z]=float(priority_vector[z])
            z+=1
        while x < No_Factors:
            Eigen_Value = Eigen_Value + (tot_factors[x] * factors_temp[x])
            x += 1
        ci = (Eigen_Value - No_Factors) / (No_Factors - 1)
        if No_Factors == 1:
            ri = 0
        elif No_Factors == 2:
            ri = 0
        elif No_Factors == 3:
            ri = 0.52
        elif No_Factors == 4:
            ri = 0.89
        elif No_Factors == 5:
            ri = 1.11
        elif No_Factors == 6:
            ri = 1.25
        elif No_Factors == 7:
            ri = 1.35
        elif No_Factors == 8:
            ri = 1.40
        elif No_Factors == 9:
            ri = 1.45
        elif No_Factors == 10:
            ri = 1.49
        elif No_Factors == 11:
            ri = 1.52
        elif No_Factors == 12:
            ri = 1.54
        elif No_Factors == 13:
            ri = 1.56
        elif No_Factors == 14:
            ri = 1.58
        elif No_Factors == 15:
            ri = 1.59
        else:
            raise sys.exit("Unable to compute due to unexpected Number of Factors")
        cr = ci / ri
        if cr < 0.10 and cr>0.0:
            str = "CR = {} < 0.10 Which is within the Acceptable Range".format(cr)
        else:
            str = "CR = {}  Which is NOT within the Acceptable Range 0.0-0.10".format(cr)

## Save the final results into the excel file in the fourth sheet under the "Final Results" name 

        fin_data = pd.DataFrame(data=[[Eigen_Value," "], [ci, " "], [ri, " "], [cr, str]], index=["Principal Eigen Value", "Consistency index", "Random Consistency Index", "Consistency Ratio"], columns=["Values", " "])
        fin_data.to_excel(writer, 'Final Results')
        writer.save()
        writer.close()
        time.sleep(3)
        os.system('start excel.exe {}'.format(opath))

## Creating and printing the results to a pcf File based on the User preference
        
        if parameters[2].value==True:
            path = opath[0:-4]+".pcf"
            pairwise_file = open(path, "w")
            pairwise_file.write(("{}".format(No_Factors)+"\n"))
            x=0
            while x< No_Factors:
                pairwise_file.write(("{}".format(factors[x])+"\n"))
                x+=1
            x = 0
            while x < No_Factors:
                j = 0
                while j <= x:
                    pairwise_file.write("{} ".format(print_table.iat[x,j]))
                    j += 1
                pairwise_file.write("\n")
                x += 1
            pairwise_file.close()
            arcpy.SetParameter(2, pairwise_file)

## Creating and printing the results to a dsf File based on the User preference

            path2 = opath[0:-4]+".dsf"
            pairwise_file2 = open(path2, "w")
            list= priority_vector.tolist()
            pairwise_file2.write(("{}".format(No_Factors)+"\n"))
            x=0
            while x< No_Factors:
                pairwise_file2.write(("{}".format(factors[x])+"\n"))
                pairwise_file2.write(("{}".format(list[x])+"\n"))
                x+=1
            pairwise_file2.close()
            arcpy.SetParameter(2, pairwise_file2)

            
        return








