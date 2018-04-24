from __future__ import print_function
import os

print("Welcome to AHP Calculator written By Khalid G. Khair")

path = raw_input("enter the location and name of the output file: ")
pairwise_file = open(path, "w")

# Decide the Number of Factors as well as their names

No_Factors = int(raw_input("Please Enter the number of factors (3-15): "))
if No_Factors > 15 or No_Factors < 3:
    p rint("The Number of Factors you have chosen can not be calculated, Exiting the program...")
    raise SystemExit
n = 0
factors = {}
com_board = [["  " for i in range(No_Factors + 1)] for z in range(No_Factors + 1)]
norm_factors = [["  " for i in range(No_Factors)] for z in range(No_Factors)]
while n < No_Factors:
    factor = raw_input("Enter factor No. {} file name: ".format(n + 1))
    factors[n] = factor
    t = 1
    com_board[0][n + 1] = factor
    com_board[n + 1][0] = factor
    com_board[n + 1][n + 1] = 1
    n += 1


# printing the comparision Table Method

def com_table(tbl):
    x = 0
    while x <= No_Factors:
        j = 0
        while j <= No_Factors:
            print("|{}|".format(tbl[x][j]), end=' ')
            j += 1
        print()
        x += 1


com_table(com_board)


# inserting the User Weights for each factors pair

def weights():
    x = 0
    while x <= No_Factors ** 2:
        z = x + 1
        while z < No_Factors:
            com_board[x + 1][z + 1] = float(raw_input("The Importance of Factor {} {} compared to Factor {} {} on a Scale of (1-9): ".format(x + 1,factors[x],z + 1,factors[z])))
            com_board[z + 1][x + 1] = 1 / (com_board[x + 1][z + 1])
            z += 1
        x += 1


weights()
com_table(com_board)

# Normalization & weight determination
x = 1
tot_factors = [0 for i in range(No_Factors)]
print("Total values for Each Column")
while x <= No_Factors:
    j = 1
    while j <= No_Factors:
        tot_factors[(x - 1)] = float(tot_factors[(x - 1)]) + com_board[j][x]
        j += 1
    print(tot_factors[(x - 1)], end=" ")
    print()
    x += 1

print("Normalized values for Each Factor")
x = 1
while x <= No_Factors:
    j = 1
    while j <= No_Factors:
        norm_factors[j - 1][x - 1] = float(com_board[j][x]) / tot_factors[x - 1]
        j += 1
    x += 1
x = 1
while x <= No_Factors:
    j = 1
    while j <= No_Factors:
        print(norm_factors[(x - 1)][j - 1], end=" ")
        j += 1
    print()
    x += 1

# Normalization & weight determination
    
print("priority vector or weight Values")
x = 0
priority_vector = [0 for i in range(No_Factors)]
while x < No_Factors:
    priority_vector[x] = sum(norm_factors[x]) / No_Factors
    print(factors[x], ": ", priority_vector[(x)])
    x += 1
    print()

##Consistancy Ratio##

x = 0
Eigen_Value = 0
while x < No_Factors:
    Eigen_Value = Eigen_Value + (tot_factors[x] * priority_vector[x])
    x += 1
ci = (Eigen_Value - No_Factors) / (No_Factors - 1)
print("Consisitancy Index: ", ci)
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
print("Random Index: ", ri)
cr = ci / ri
if cr < 0.10:
    print("CR = {} < 0.10 Which is within the Acceptable Range".format(cr))
else:
    print("CR = {} > 0.10 Which is NOT within the Acceptable Range".format(cr))

# Printing to File

pairwise_file.write((str(No_Factors)+"\n"))
x=0
while x< No_Factors:
    pairwise_file.write((str(factors[x])+"\n"))
    x+=1

x = 1
while x <= No_Factors :
    z =  1
    while z <= x:
        pairwise_file.write((str(com_board[x][z])+" "))
        z += 1
    pairwise_file.write("\n")
    x += 1
pairwise_file.close()
