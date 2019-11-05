# xlsComparer
tool which does a diff between two xls files and highlights similarities

takes as inputs:
- older file (xlsx)
- newer file (xlsx)
- output filename

xlsComparer will read in the older file, generate a concatenated string for each row and insert it into a map. Then it opens the newer file and copies it to the new output filename such that the rows which have been seen already will be marked in a green background.
