'Create a Loop that runs through all of the stocks for each quarter and outputs the following information:
    ' The Ticker Symbol - place this in a new column two to the right of the last column with values
    ' Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
    ' The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
    ' The total stock volume of the stock


    This VBA script uses stock data from 4 quarters, establishes each worksheet in the workbook so the subroutine can loop through all workbooks
    The first and last row are established so ranges aren't limited, because each data set is a different length
    It adds column titles for the requested returned data in each of the worksheets
    The For loop is comparing the ticker names in each row and storing Total Volume, Open Price and Close Price for each of the ticker
    Formulas then be executed agains the stored Open and Close Price before being placed in the corresponding columns of the summary table
    Formatting is happening along the way, in addition to conditional formatting and auto fitting column width to fit the placed values
    
