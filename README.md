# HouseMaker


This programme will do below,

    read the original data excel
    check and report total cartons need to split
    check and report total number of unique items
    check and report total number of unique HS CODE
    report total pcs, ct, duty and vat for each item
    group and check HS CODE, mark them with Safe and Sensitive with criteria
    randomly split total cartons into houses based on user input
    group into cartons to two category, one with single SKU per FBA, another with multi SKU per FBA
    assign the cartons to houses according the numbers generated
    save the houses to excel file, each house on one sheet.

to add features

    calculate DUTY and VAT for each house
    add input and output location
    format output excel
    check if GW is greater than NW
    create output folder, add time stamp on output file
    create a log file for the program
    create another pivot talbe for each sheet on excel, calculate duty vat etc


29/01/2019
  revised version according to new template, file name SplitInvoiceGui_newTemplate.py

  pre processing steps (need to be done before running the program)

  a. delete unwanted lines at the end of excel file
  b. unmerge 'Marks & No' column and refill the value, unmerge 'CTN' column but don't fill the value. use below VBA CODE

                  Sub UnMergeFill()

                  Dim cell As Range, joinedCells As Range, wsh As Worksheet

                      For Each wsh In ThisWorkbook.Worksheets
                          For Each cell In wsh.UsedRange
                              If Not cell.Column = 17 Then   'ignore column Q
                                  If cell.MergeCells Then
                                      Set joinedCells = cell.MergeArea
                                      cell.MergeCells = False
                                      joinedCells.Value = cell.Value
                                  End If
                              Else   'if in column Q, do unmerge only
                                  If cell.MergeCells Then
                                      Set joinedCells = cell.MergeArea
                                      cell.MergeCells = False
                                  End If
                              End If
                          Next
                      Next

                  End Sub
