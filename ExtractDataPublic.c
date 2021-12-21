#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "../include/xls.h" // for working with Excel worksheets

#define charLength 2048
#define defaultPath "C:/Users/[user]/Downloads/"
#define defaultName "Schedule.xls"

int checkifempty(char * str);

int main(int argc, char *argv[]) 
{
    // Excel variables
    xlsWorkBook * pWB;
    xlsWorkSheet * pWS;
    // Weekdays
    char * weekDays[14] = { "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Monday", "Tuesday" };

    // Path to Excel file
    char sheetPath[charLength] = { 0 };
    strncpy(sheetPath, defaultPath, charLength);
    // if user-entered name
    if (argc == 3) strncat(sheetPath, argv[2], charLength - strlen(defaultPath));
    // else assuming default name of Schedule
    else strncat(sheetPath, defaultName, charLength - strlen(defaultPath));
    
    // Sheet number
    int sheetNumber = 2;
    if (argc >= 2) sheetNumber += atoi(argv[1]);

    xlsRow * row;
    xlsCell * cell;
    WORD t, tt;
    xls_error_t code = LIBXLS_OK;
    xls(10);

    // Open Excel file      
    pWB = xls_open_file(sheetPath, "UTF-8", &code);
    if (pWB == NULL) {
        printf("Error: WorkBook is null! Exiting...\n");
        printf("libxls error:\n%s\n", xls_getError(code));
        return 1;
    }
    else printf("Passed: WorkBook existence check\n");

    // get most recent sheet
    pWS = xls_getWorkSheet(pWB, pWB->sheets.count-sheetNumber); 
    if (pWS == NULL) {
        printf("Error: WorkSheet is null! Exiting...\n");
        printf("libxls:\n%s\n", xls_getError(code));
        return 1;
    }
    else printf("Passed: WorkSheet existence check\n");

    // Parse worksheet picked
    if((code = xls_parseWorkSheet(pWS)) != LIBXLS_OK) {
        printf("Error: WorkSheet could not be parsed correctly! Exiting...\n");
        printf("libxls error:\n%s\n", xls_getError(code));
        return 1;
    }
    else printf("Passed: WorkSheet parse check\n");

    printf("\n\n\nWork Schedule\n\n");
    // go through Excel sheet rows and columns
    for (t = 1; t < 60; t++) // t <= pWS->rows.lastrow
    {
        row = &pWS->rows.row[t];

        if (row->cells.cell[0].str && !strcmp(row->cells.cell[0].str, "[name]"))
        {
            for (tt = 1; tt < 15; tt++) // tt <= pWS->rows.lastcol
            {
                cell = &row->cells.cell[tt];
                if (cell->id != XLS_RECORD_NUMBER) 
                {
                    if (cell->str != NULL && cell->str[0] != '\0' && !checkifempty(cell->str)) 
                    {
                        if (!strcmp(cell->str, "OFF")) printf("[name] has %s OFF\n", weekDays[(tt-1)/2]);
                        else printf("[name] works %s from %s\n", weekDays[(tt-1)/2], cell->str);
                    }
                    else printf("[name] does not work %s\n", weekDays[(tt-1)/2]);
                }
             }

            printf("\n\n");
            return 0;
        }
    }

    return 2;
}

int checkifempty(char * str)
{
    int i = 0, spaces = 0;

    for (; str[i] != '\0'; i++) {
        if (isspace(str[i])) spaces++;
    }

    if (i == spaces) return 1;
    else return 0;
}

/*
Compiling and generating:
gcc ExtractData.c -I[path to include dir] -lxlsreader -o extract

Running this program:
./extract [number of weeks ago*] [sheet name**]
* optional if not using third variable
** optional
*/