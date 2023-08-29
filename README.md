# Automatic YouTube Promotion Metrics

A seamless way to visualize and interpret your YouTube Promotion data while it is currently in Beta. Simply copy, paste, and let the program generate a well-organized Excel sheet, detailing key metrics such as 'Cost per Sub' and 'Impression to Sub Ratio' as well as auto-formatting and intelligently sorting data based on performance metrics.


Additional Features:
- Promotion lists are separated by Status
- Total Active and Total counts are shown
- Rows are alternatively colored for easy readability
- Headings and List Titles are color coded for easy readability

![Sample Output](https://github.com/Ranchy101/YouTube-Promotion-Metrics/assets/42690717/e4e24120-4c2f-43ea-8a06-dedbae4f1b90)

Instructions:
1. Open your YouTube Promotions tab.
2. Scroll down to bottom and set Rows per Page to the maximum.
3. Copy and paste everything starting from "Promotion" all the way down to "Rows per page..."
4. Paste into promotions.txt file
5. Make sure previous output.xlsx are closed and that promotions.txt is saved in same directory
6. Compile and Run TestToExcel.java

Changelog:
- Automatic removal of "trash" rows in promotions.txt file

Upcoming Changes:
- An additional sorted list below the current one that sorts by Cost per Sub for ALL promotions.
- Impression to View Ratio column
- View to Sub Ratio column

End Goal:
Visualization in the form of bar graphs across multiple days to measure if a promotion is getting better or worse over time. Would require multiple output.xlsx files organized/named by date. 
For example, the program would take 8-27-23.xlsx, 8-28-23.xlsx, and 8-29-23.xlsx; and create a seperate graph.xlsx file showing 2 bar graphs for every promotion that was active during all 3 files. One bar graph for Cost per Sub and another for Impression to Sub Ratio.
  
