import pandas as pd
import numpy as np
from pprint import pprint
import sys
import time
import datetime
import os
from argparse import ArgumentParser
from gooey import Gooey, GooeyParser
import openpyxl
import json

@Gooey(program_name="Creat Splits from Combined Data")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """

    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Create Splits')
    parser.add_argument('Source_ExcelFile',
                        action='store',
                        default=stored_args.get('Source_ExcelFile'),
                        widget='FileChooser',
                        help="Source Excel files with combined data",
                        )
    parser.add_argument('output_folder',
                        action='store',
                        default=stored_args.get('output_folder'),
                        help="Output directory to save summary report",
                        )
    parser.add_argument('number_of_splits',
                        action='store',
                        default=stored_args.get('number_of_splits'),
                        help='Number of splits you would like to create',
                        )
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as dictionary
        json.dump(vars(args), data_file)
    return args






# def welcome():
#     originFile = input('Please enter the full path of the source excel file (ending with .xlsx): \n')
#
#     destFolder = input('Please enter the folder name you want to save the gererated splits: \n')
#
#     return originFile, destFolder

def SplitMaker(filename, destFolder, numberOfSplits):


#    filename, destFolder = welcome()
#    nos = input('please enter the number of splits you would like to create\n')      #input number of splits you want
    nos = int(numberOfSplits)


    timestr = time.strftime("%Y%m%d-%H%M")
    destFolder = destFolder + '_' + timestr
    savePath = os.path.join(os.environ['USERPROFILE'], 'Documents', 'SplitInvoices', destFolder)

    if os.path.exists(savePath) == False:
        os.makedirs(savePath)
    os.chdir(savePath)




    ############# save the output to a txt file
    timestr = time.strftime("%Y%m%d-%H%M")
    old_stdout = sys.stdout

    log_file = open('report_' + timestr + '.txt', 'w')

    sys.stdout = log_file

    print('Sarting program at: {}'.format(datetime.datetime.now().strftime("%I:%M%p on %B %d, %Y")))
    #############

    # input main part program below

    df = pd.read_excel(filename)
    df['HS'] = df['HS'].apply(str)  #convert HS CODE column from int to str


    def check_HS(x):
        if x.startswith('42') | x.startswith('61') | x.startswith('62'):
            return 'Sensitive'
        else:
            return 'Safe'

    df['Sensitivity']= df['HS'].apply(check_HS)    # add a columns to check if HS is safe or Sensitive

    df['Sensitivity'] =df['Sensitivity'].astype('category') # set Sensitivity column as category

    df['CT_PERC'] = np.nan
    df['QTY_PERC'] = np.nan

    print(df.pivot_table(index=['Sensitivity'], values=['CT', 'QTY'], aggfunc=[np.sum]))


    sens_ct = df.loc[df['Sensitivity'] == 'Sensitive'].CT.sum()     #sens_ct means total CARTONS of sensitivity HS CODE
    total_ct = df.CT.sum()
    sens_ct_perc = sens_ct / total_ct
    print('Sensitive HS CODE carton percentage: {:2.2%}'.format(sens_ct_perc))


    sens_pc = df.loc[df['Sensitivity'] == 'Sensitive'].QTY.sum()    # sens_pc means total PCS of sensitivity HS CODE
    total_qty = df.QTY.sum()
    sens_pc_perc = sens_pc / total_qty
    print('Sensitive HS CODE pcs percentage: {:2.2%}'.format(sens_pc_perc))

    print('\n')


    columns = df.columns
    # print(columns)

    # num_unique = df.groupby('MARKS')['MARKS'].nunique()  # count unique FBA number, which equls to how many cartons
    # print('Total cartons of unique FBA: {} ctns'.format(len(num_unique)))

    nof = len(df.groupby('MARKS')) # number of cartons
    print('Total cartons of unique FBA: {} ctns'.format(nof))

    noi = len(df.groupby('Revised Name')) # number of items
    print('Total numbers of unique items: {}'.format(noi))

    noHS = len(df.groupby('HS'))   #number of unique HS CODE
    print('Total numbers of unique HS CODE: {}'.format(noHS))
    l_HS = list(df.drop_duplicates('HS', keep='first')['HS'])      #list of unique HS CODE
    pprint(l_HS)

    # Below code will filter the items with sensitive HS CODE, which starts with 42, 61 and 62, nos_HS, number of sensitive HS CODE
    nos_HS = len(df[df['HS'].str.startswith('42') | df['HS'].str.startswith('61') | df['HS'].str.startswith('62')])
    print('Total lines of sensitive HS CODE, start with 42, 61, 62 : {}'.format(nos_HS))

    table = pd.pivot_table(df,index=['Revised Name'], values=['QTY', 'CT', 'DUTY', 'VAT'], aggfunc=[np.sum], margins=True, margins_name='Total')
    column_order = ['QTY', 'CT', 'DUTY', 'VAT']
    table2 = table.reindex(column_order, level=1, axis=1)   #pivot table is multiindex DATAFRAME, need to use level argument when reindex
    table2

    Duty_total = df['DUTY'].sum()            # print out total DUTY and VAT
    print('DUTY total: {:.2f}'.format(Duty_total))
    Vat_total = df['VAT'].sum()
    print('VAT total: {:.2f}\n'.format(Vat_total))

    for name, group in df.groupby('Revised Name'):
        print('{0} \n has {1} lines \n   this item is {2}'.format(name, len(group), group.Sensitivity.iloc[0]))
        print('*' * 40)

    # call split_cartons function to create a list, pass variable nof
    def split_cartons(tc, nos):
        import random

        #todo, add another parameter base number of cartons in 1 split, with default value of 12
        # tc ---- total cartons in MAWB, tc need to be larger than 12, otherwise will get division 0 error
    #    nos = tc // 12            # nos --- number of splits we'll create
        ls = []        # ls ---- list of splits that contains number of cartons of each split

        for x in range(nos):
            ls.append(random.randint(12, 14))   # generate a number between 9 to 15, assign to ls



        diff = tc - sum(ls)                # workout the difference between tc and sum(ls)


        if diff == 0:
            pass
        elif diff > 0:                            # equally distribut the differnce to the splits
            if diff // len(ls) > 0:
                ls = [k + diff // len(ls) for k in ls]
            for x in range(diff % len(ls) ):
                ls[x] += 1
        else:
            diff = -diff
            if diff // len(ls) > 0:
                ls = [k - diff // len(ls) for k in ls]
            for x in range(diff % len(ls) ):
                ls[x] -= 1

        print('Total {} cartons in MAWB'.format(tc))
        print('split into {} splits, total {} cartons. details as below'.format(len(ls), sum(ls)))
        for i in range(len(ls)):
            print('split {}: {} cartons'.format(i + 1, ls[i]))
        return ls


    split_result = split_cartons(nof, nos)

    # print(len(split_result))

    df_multiSKU = df.loc[df['MARKS'].duplicated(keep=False), :]  # create dataframe for cartons that 1 FBA contains multi SKU
    df_multiSKU['MARKS'].unique()
    qouf = len(df_multiSKU['MARKS'].unique())         # cartons of FBA contains multi SKU

    df_singleSKU = df.drop_duplicates('MARKS', keep=False) # create dataframe for cartons that 1 FBA contains 1 SKU
    df_singleSKU = df_singleSKU.sort_values(by='Sensitivity', ascending=False) # sort the order of dataframe by Sensitivity, put sensitive in the beginning


    #create a list of empty DataFrame with number of splits
    x = len(split_result)
    lst = []
    for i in range(x):
    #    lst.append(pd.DataFrame(index=range(result[i])))  #create a list of empty DataFrame with fixed rows
        lst.append(pd.DataFrame())   #create a list of empty DataFrame


    # assign items into splits

    split_limit = len(split_result)

    counter = 0

    for name, group in df_singleSKU.groupby('Sensitivity', sort=False):   #try put the sensitivity item at the beginning

        for index, row in group.iterrows():

            if len(lst[counter%split_limit]) < split_result[counter%split_limit]:
                lst[counter%split_limit] = lst[counter%split_limit].append(row, ignore_index=False, sort=True)
            else:
                for i in range(split_limit):

                    if len(lst[(counter + i + 1)%split_limit]) < split_result[(counter + i + 1)%split_limit]:
                        lst[(counter + i + 1)%split_limit] = lst[(counter + i + 1)%split_limit].append(row, ignore_index=False, sort=True)
                        break
            counter += 1

    # carry on add multi SKU FBAS
    for name, group in df_multiSKU.groupby('MARKS'):
        if lst[counter%split_limit].CT.sum() < split_result[counter%split_limit]:
            lst[counter%split_limit] = lst[counter%split_limit].append(group, ignore_index=False, sort=True)
        else:
            for i in range(split_limit):
                if lst[(counter + i + 1)%split_limit].CT.sum() < split_result[(counter + i + 1)%split_limit]:
                    lst[(counter + i + 1)%split_limit] = lst[(counter + i + 1)%split_limit].append(group, ignore_index=False, sort=True)
                    break
        counter += 1

    print('\n')
    print('*' * 40)
    print('HS percentage for whole shipment as below')
    print('Sensitive HS CODE carton percentage: {:2.2%}'.format(sens_ct_perc))
    print('Sensitive HS CODE pcs percentage: {:2.2%}'.format(sens_pc_perc))
    print('*' * 40)
    print('\n')



    # add two columns, CT PERCENTAGE AND QTY PERCENTAGE
    for index, data in enumerate(lst):
        data['CT_PERC'] = round(data['CT'] / data['CT'].sum() * 100, 2)
        data['QTY_PERC'] = round(data['QTY'] / data['QTY'].sum() * 100, 2)
        print('\n')
        print('*' * 40)
        print('split {} : {:.0f} ctns'.format(index + 1, data.CT.sum()))
        print(data.pivot_table(index=['Sensitivity'], values=['CT_PERC', 'QTY_PERC'], aggfunc=np.sum))
        print('*' * 40)

    #save the data result to excel
    output_file_name = 'output_' + str(nos) + 'sp_' + timestr + '.xlsx'
    writer = pd.ExcelWriter(output_file_name)
    for data, i in zip(lst, range(len(split_result))):
        data.to_excel(writer, sheet_name = 'split_'+str(i+1) +'--'+ str(split_result[i]) + '_pcs', columns=columns)
    writer.save()




    #finishing part, will close the log file
    print('Closing program at: {}'.format(datetime.datetime.now().strftime("%I:%M%p on %B %d, %Y")))

    sys.stdout = old_stdout

    log_file.close()


    print('{} splits created sucessfully.'.format(nos))




if __name__ == '__main__':
    conf = parse_args()
    SplitMaker(conf.Source_ExcelFile, conf.output_folder, conf.number_of_splits)
